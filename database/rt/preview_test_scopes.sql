-- SOMENTE LEITURA: plano exato dos escopos iniciais do executor.
-- NAO INSERTAR antes de --run-fixtures: setUpClass exige tabelas RT vazias.
WITH planned(principal,company_id,condominium_id,period,purpose) AS (VALUES
 ('rt_test_a','10000000-0000-0000-0000-000000000001'::uuid,
  '30000000-0000-0000-0000-000000000003'::uuid,DATE '2026-10-01','A / condominio 1'),
 ('rt_test_a','10000000-0000-0000-0000-000000000001'::uuid,
  '40000000-0000-0000-0000-000000000004'::uuid,DATE '2026-10-01','A / condominio 2'),
 ('rt_test_a','10000000-0000-0000-0000-000000000001'::uuid,
  '50000000-0000-0000-0000-000000000005'::uuid,DATE '2026-10-01','A / recuperacao'),
 ('rt_test_b','20000000-0000-0000-0000-000000000002'::uuid,
  '30000000-0000-0000-0000-000000000003'::uuid,DATE '2026-10-01','B / condominio 1')
)
SELECT * FROM planned ORDER BY principal,condominium_id;
-- rt_test_none: ZERO escopos. Cada cenario cria ainda um condominio UUID novo
-- para rt_test_a, no mesmo periodo, pelo administrador de fixtures.
