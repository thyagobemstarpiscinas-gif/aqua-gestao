-- Etapas 1 e 2. Executar somente em banco vazio exclusivo de homologacao.
-- Migracao unica: uma segunda aplicacao falha sem modificar a instalacao.
BEGIN;
CREATE SCHEMA rt_persist;
REVOKE ALL ON SCHEMA rt_persist FROM PUBLIC;
CREATE ROLE rt_persist_client NOLOGIN NOSUPERUSER NOCREATEDB NOCREATEROLE NOINHERIT NOBYPASSRLS;

CREATE TABLE rt_persist.schema_migrations (
    version integer PRIMARY KEY,
    installed_at timestamptz NOT NULL DEFAULT clock_timestamp()
);
INSERT INTO rt_persist.schema_migrations(version) VALUES (1);

CREATE TABLE rt_persist.companies (
    company_id uuid PRIMARY KEY,
    label text NOT NULL
);
CREATE TABLE rt_persist.condominiums (
    company_id uuid NOT NULL REFERENCES rt_persist.companies,
    condominium_id uuid NOT NULL,
    label text NOT NULL,
    PRIMARY KEY (company_id, condominium_id)
);
-- Provisionamento administrativo. Nunca recebe contexto escolhido pelo navegador.
-- session_user e a identidade da conexao PostgreSQL, inclusive dentro de SECURITY DEFINER.
CREATE TABLE rt_persist.access_scopes (
    principal name NOT NULL,
    company_id uuid NOT NULL,
    condominium_id uuid NOT NULL,
    period date NOT NULL CHECK (extract(day FROM period) = 1),
    PRIMARY KEY (principal, company_id, condominium_id, period),
    FOREIGN KEY (company_id, condominium_id)
        REFERENCES rt_persist.condominiums
);
CREATE TABLE rt_persist.documents (
    company_id uuid NOT NULL,
    condominium_id uuid NOT NULL,
    period date NOT NULL CHECK (extract(day FROM period) = 1),
    head bigint NOT NULL DEFAULT 0 CHECK (head >= 0),
    PRIMARY KEY (company_id, condominium_id, period),
    FOREIGN KEY (company_id, condominium_id)
        REFERENCES rt_persist.condominiums
);
CREATE TABLE rt_persist.revisions (
    company_id uuid NOT NULL,
    condominium_id uuid NOT NULL,
    period date NOT NULL,
    revision bigint NOT NULL CHECK (revision > 0),
    base_revision bigint NOT NULL CHECK (base_revision = revision - 1),
    restored_from bigint,
    schema_version integer NOT NULL CHECK (schema_version = 1),
    payload jsonb NOT NULL CHECK (jsonb_typeof(payload) = 'object'),
    content_hash text NOT NULL,
    author name NOT NULL,
    created_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    PRIMARY KEY (company_id, condominium_id, period, revision),
    FOREIGN KEY (company_id, condominium_id, period)
        REFERENCES rt_persist.documents,
    FOREIGN KEY (company_id, condominium_id, period, restored_from)
        REFERENCES rt_persist.revisions(company_id, condominium_id, period, revision)
);
CREATE TABLE rt_persist.save_requests (
    company_id uuid NOT NULL,
    condominium_id uuid NOT NULL,
    period date NOT NULL,
    request_id uuid NOT NULL,
    base_revision bigint NOT NULL CHECK (base_revision >= 0),
    restored_from bigint,
    payload jsonb NOT NULL,
    author name NOT NULL,
    outcome text NOT NULL CHECK (outcome IN ('saved', 'conflict')),
    observed_head bigint NOT NULL,
    result jsonb NOT NULL,
    created_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    PRIMARY KEY (company_id, condominium_id, period, request_id),
    FOREIGN KEY (company_id, condominium_id, period)
        REFERENCES rt_persist.documents,
    FOREIGN KEY (company_id, condominium_id, period, restored_from)
        REFERENCES rt_persist.revisions(company_id, condominium_id, period, revision)
);
-- Apenas registro transacional; nenhum exportador implementado nestas etapas.
CREATE TABLE rt_persist.outbox (
    company_id uuid NOT NULL,
    condominium_id uuid NOT NULL,
    period date NOT NULL,
    revision bigint NOT NULL,
    state text NOT NULL DEFAULT 'pending'
        CHECK (state IN ('pending', 'sending', 'retry', 'confirmed')),
    created_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    PRIMARY KEY (company_id, condominium_id, period, revision),
    FOREIGN KEY (company_id, condominium_id, period, revision)
        REFERENCES rt_persist.revisions
);
CREATE FUNCTION rt_persist.reject_mutation() RETURNS trigger
LANGUAGE plpgsql SET search_path = pg_catalog, rt_persist AS $$
BEGIN
    RAISE EXCEPTION 'immutable record' USING ERRCODE = '55000';
END;
$$;
CREATE TRIGGER revisions_immutable BEFORE UPDATE OR DELETE ON rt_persist.revisions
FOR EACH ROW EXECUTE FUNCTION rt_persist.reject_mutation();
CREATE TRIGGER requests_immutable BEFORE UPDATE OR DELETE ON rt_persist.save_requests
FOR EACH ROW EXECUTE FUNCTION rt_persist.reject_mutation();

CREATE FUNCTION rt_persist.assert_scope(c uuid, d uuid, p date) RETURNS void
LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, rt_persist AS $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM rt_persist.access_scopes
        WHERE principal = session_user AND company_id = c
          AND condominium_id = d AND period = p
    ) THEN
        RAISE EXCEPTION 'scope denied' USING ERRCODE = '42501';
    END IF;
END;
$$;

CREATE FUNCTION rt_persist.save(
    c uuid, d uuid, p date, operation uuid, base bigint,
    body jsonb, restore_revision bigint DEFAULT NULL
) RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, rt_persist AS $$
DECLARE
    h bigint;
    prior rt_persist.save_requests%ROWTYPE;
    answer jsonb;
BEGIN
    PERFORM rt_persist.assert_scope(c, d, p);
    IF operation IS NULL OR base IS NULL OR base < 0
       OR body IS NULL OR jsonb_typeof(body) <> 'object'
       OR jsonb_typeof(body->'campos') IS DISTINCT FROM 'object'
       OR octet_length(body::text) > 8388608 THEN
        RAISE EXCEPTION 'invalid save contract' USING ERRCODE = '22023';
    END IF;
    -- Identidade pertence ao envelope relacional. Metadados legados nao podem
    -- contradize-lo; campos de exibicao nao sao usados para autorizar acesso.
    IF body ? 'identity' AND body->'identity' IS DISTINCT FROM jsonb_build_object(
        'company_id', c, 'condominium_id', d, 'period', p
    ) THEN
        RAISE EXCEPTION 'identity mismatch' USING ERRCODE = '22023';
    END IF;
    INSERT INTO rt_persist.documents(company_id, condominium_id, period)
    VALUES (c, d, p) ON CONFLICT DO NOTHING;
    SELECT head INTO h FROM rt_persist.documents
    WHERE company_id = c AND condominium_id = d AND period = p FOR UPDATE;
    -- O lock tambem serializa repeticoes simultaneas do mesmo operation.
    SELECT * INTO prior FROM rt_persist.save_requests
    WHERE company_id = c AND condominium_id = d AND period = p
      AND request_id = operation;
    IF FOUND THEN
        IF prior.base_revision <> base OR prior.payload IS DISTINCT FROM body
           OR prior.restored_from IS DISTINCT FROM restore_revision
           OR prior.author <> session_user THEN
            RAISE EXCEPTION 'idempotency key reused with different request'
                USING ERRCODE = '22023';
        END IF;
        RETURN prior.result;
    END IF;
    IF base > h THEN
        RAISE EXCEPTION 'base revision does not exist' USING ERRCODE = '22023';
    END IF;
    IF restore_revision IS NOT NULL AND NOT EXISTS (
        SELECT 1 FROM rt_persist.revisions
        WHERE company_id = c AND condominium_id = d AND period = p
          AND revision = restore_revision
    ) THEN
        RAISE EXCEPTION 'restore revision outside document or absent'
            USING ERRCODE = '22023';
    END IF;
    IF base <> h THEN
        answer := jsonb_build_object('status', 'conflict', 'request_id', operation,
                                      'observed_head', h);
        INSERT INTO rt_persist.save_requests
            (company_id, condominium_id, period, request_id, base_revision,
             restored_from, payload, author, outcome, observed_head, result)
        VALUES (c, d, p, operation, base, restore_revision, body, session_user,
                'conflict', h, answer);
        RETURN answer; -- retorno normal: o chamador deve COMMIT para preservar conflito
    END IF;
    INSERT INTO rt_persist.revisions
        (company_id, condominium_id, period, revision, base_revision,
         restored_from, schema_version, payload, content_hash, author)
    VALUES (c, d, p, h + 1, h, restore_revision, 1, body,
            encode(sha256(convert_to(body::text, 'UTF8')), 'hex'), session_user);
    UPDATE rt_persist.documents SET head = h + 1
    WHERE company_id = c AND condominium_id = d AND period = p;
    INSERT INTO rt_persist.outbox(company_id, condominium_id, period, revision)
    VALUES (c, d, p, h + 1);
    answer := jsonb_build_object('status', 'saved', 'request_id', operation,
                                 'revision', h + 1, 'sync_state', 'pending');
    INSERT INTO rt_persist.save_requests
        (company_id, condominium_id, period, request_id, base_revision,
         restored_from, payload, author, outcome, observed_head, result)
    VALUES (c, d, p, operation, base, restore_revision, body, session_user,
            'saved', h, answer);
    RETURN answer;
END;
$$;

CREATE FUNCTION rt_persist.history(c uuid, d uuid, p date) RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, rt_persist AS $$
DECLARE answer jsonb;
BEGIN
    PERFORM rt_persist.assert_scope(c, d, p);
    SELECT coalesce(jsonb_agg(to_jsonb(r) ORDER BY revision), '[]'::jsonb)
    INTO answer FROM rt_persist.revisions r
    WHERE company_id = c AND condominium_id = d AND period = p;
    RETURN answer;
END;
$$;
CREATE FUNCTION rt_persist.conflicts(c uuid, d uuid, p date) RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, rt_persist AS $$
DECLARE answer jsonb;
BEGIN
    PERFORM rt_persist.assert_scope(c, d, p);
    SELECT coalesce(jsonb_agg(to_jsonb(r) ORDER BY created_at, request_id), '[]'::jsonb)
    INTO answer FROM rt_persist.save_requests r
    WHERE company_id = c AND condominium_id = d AND period = p AND outcome = 'conflict';
    RETURN answer;
END;
$$;
-- PUBLIC nao recebe tabelas nem funcoes, inclusive via Supabase Data API.
REVOKE ALL ON ALL TABLES IN SCHEMA rt_persist FROM PUBLIC, rt_persist_client;
REVOKE ALL ON ALL FUNCTIONS IN SCHEMA rt_persist FROM PUBLIC, rt_persist_client;
GRANT USAGE ON SCHEMA rt_persist TO rt_persist_client;
GRANT EXECUTE ON FUNCTION rt_persist.save(uuid, uuid, date, uuid, bigint, jsonb, bigint)
    TO rt_persist_client;
GRANT EXECUTE ON FUNCTION rt_persist.history(uuid, uuid, date),
    rt_persist.conflicts(uuid, uuid, date) TO rt_persist_client;
ALTER DEFAULT PRIVILEGES IN SCHEMA rt_persist REVOKE EXECUTE ON FUNCTIONS FROM PUBLIC;
COMMIT;
