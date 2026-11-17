"""Escolha de canais, envio SMTP explícito e compartilhamento local de documentos."""
import base64
import hashlib
import json
from pathlib import Path
import tempfile
from urllib.parse import quote

from utils.email_confirmacao import mensagem_com_confirmacao


CANAIS = ("E-mail", "WhatsApp", "Ambos")
MIMES = {".pdf": "application/pdf", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}


def preparar_arquivos(caminhos, uploads):
    """Snapshot dos anexos escolhidos; falha sem enviar se algum estiver ausente."""
    arquivos = []
    for item in [*caminhos, *uploads]:
        nome = Path(str(item.name).replace("\\", "/")).name
        extensao = Path(nome).suffix.lower()
        if extensao not in MIMES:
            raise ValueError("Selecione apenas documentos PDF ou DOCX.")
        try:
            conteudo = bytes(item.getbuffer()) if hasattr(item, "getbuffer") else Path(item).read_bytes()
        except OSError:
            raise ValueError(f"O arquivo {nome} não está mais disponível. Selecione novamente.") from None
        if not conteudo:
            raise ValueError(f"O arquivo {nome} está vazio.")
        arquivos.append({"nome": nome, "mime": MIMES[extensao], "conteudo": conteudo})
    if not arquivos:
        raise ValueError("Selecione ou adicione pelo menos um documento.")
    return arquivos


def executar_envio(canal, arquivos, enviar_email, *, destinatario, assunto, mensagem, cc="", bcc=""):
    """Não chama SMTP para WhatsApp; não repete falhas e mantém resultados separados."""
    if canal not in CANAIS:
        raise ValueError("Canal de envio inválido.")
    resultado = {"email": None, "whatsapp_preparado": canal in ("WhatsApp", "Ambos")}
    if canal in ("E-mail", "Ambos"):
        if not destinatario.strip() or not assunto.strip():
            raise ValueError("Informe o destinatário e o assunto do e-mail.")
        with tempfile.TemporaryDirectory(prefix="acqua_envio_") as tmp:
            anexos = []
            for i, arquivo in enumerate(arquivos):
                # Diretórios distintos preservam nomes iguais sem sobrescrever anexos.
                pasta = Path(tmp) / str(i)
                pasta.mkdir()
                caminho = pasta / arquivo["nome"]
                caminho.write_bytes(arquivo["conteudo"])
                anexos.append(caminho)
            resultado["email"] = enviar_email(destinatario, assunto, mensagem, anexos, cc=cc, bcc=bcc)
    return resultado


def _identidade_selecao(contexto, caminhos, uploads):
    detalhes = []
    for caminho in caminhos:
        p = Path(caminho)
        try:
            info = p.stat()
            detalhes.append((str(p.resolve()), info.st_mtime_ns, info.st_size))
        except OSError:
            detalhes.append((str(p), "ausente"))
    for upload in uploads:
        detalhes.append((upload.name, hashlib.sha256(upload.getbuffer()).hexdigest()))
    return hashlib.sha256(json.dumps([contexto, detalhes], ensure_ascii=False).encode()).hexdigest()


def exibir_compartilhamento(st, arquivos, *, mensagem, titulo, telefone, key):
    from streamlit.components.v2 import component

    st.info("WhatsApp: documentos preparados. Toque em compartilhar, escolha o WhatsApp e o contato, e confirme o envio no aparelho.")
    st.caption("O compartilhamento depende do navegador e dos aplicativos instalados. O app não confirma a entrega pelo WhatsApp.")
    compartilhar = component(
        "acqua_compartilhar_documentos",
        html='<button type="button">Compartilhar documentos no celular</button><p role="status" aria-live="polite"></p>',
        css='button{background:#177d53;color:white;border:0;border-radius:8px;padding:12px 16px;font:600 15px sans-serif;cursor:pointer;width:100%}\nbutton:disabled{opacity:.55;cursor:default}\np{font:14px/1.5 sans-serif;color:var(--st-text-color)}',
        js=Path(__file__).with_name("compartilhar_documentos.js").read_text(),
    )
    compartilhar(key=key + "_compartilhar", data={
        "titulo": titulo, "mensagem": mensagem,
        "arquivos": [{"nome": a["nome"], "mime": a["mime"], "base64": base64.b64encode(a["conteudo"]).decode("ascii")} for a in arquivos],
    })
    with st.expander("Alternativa: baixar e anexar no WhatsApp", expanded=False):
        for i, arquivo in enumerate(arquivos):
            st.download_button("Baixar " + arquivo["nome"], arquivo["conteudo"], file_name=arquivo["nome"], mime=arquivo["mime"], key=f"{key}_baixar_{i}", on_click="ignore")
        numero = "".join(c for c in telefone if c.isdigit())
        if numero and not numero.startswith("55"):
            numero = "55" + numero
        if numero:
            st.link_button("Abrir conversa com a mensagem", f"https://wa.me/{numero}?text={quote(mensagem)}")
            st.caption("Esse link abre apenas a mensagem. Anexe os documentos baixados antes de enviar.")


def exibir_envio_documentos(st, *, nome_condominio, documentos, email_cliente, mensagem_padrao,
                           key, enviar_email, email_configurado, whatsapp_cliente="", selecionados_padrao=None):
    """Formulário único nos relatórios e na central; reruns nunca disparam envio."""
    key += "_" + hashlib.sha256(nome_condominio.strip().casefold().encode()).hexdigest()[:12]
    canal = st.radio("Enviar por", CANAIS, horizontal=True, key=key + "_canal")
    tem_email = canal in ("E-mail", "Ambos")
    tem_whatsapp = canal in ("WhatsApp", "Ambos")
    destinatario = assunto = cc = bcc = telefone = ""
    cfg_ok, cfg_erro = email_configurado
    if tem_email:
        c1, c2 = st.columns(2)
        with c1:
            destinatario = st.text_input("Destinatário do e-mail", value=email_cliente, key=key + "_destinatario")
        with c2:
            assunto = st.text_input("Assunto", value=f"Documentação técnica ACQUA GESTÃO - {nome_condominio}", key=key + "_assunto")
        with st.expander("CC / CCO"):
            cc = st.text_input("CC", key=key + "_cc")
            bcc = st.text_input("CCO", key=key + "_bcc")
        if not cfg_ok:
            st.warning("Envio de e-mail indisponível: " + cfg_erro)
    if tem_whatsapp:
        telefone = st.text_input("WhatsApp do cliente (opcional, para abrir a conversa)", value=whatsapp_cliente, key=key + "_telefone")
        st.caption("Ao compartilhar o PDF no celular, você escolhe o WhatsApp e o contato na tela do aparelho.")
    mensagem = st.text_area("Mensagem", value=mensagem_padrao, height=180, key=key + "_mensagem")
    if tem_email:
        st.caption("O e-mail inclui automaticamente o pedido de confirmação de recebimento e a assinatura profissional da ACQUA.")
        with st.expander("Ver texto do e-mail com pedido de confirmação"):
            st.text(mensagem_com_confirmacao(mensagem))
    documentos = [Path(p) for p in documentos]
    # IDs distintos evitam misturar documentos de mesmo nome em pastas diferentes.
    mapa = {str(p.resolve()): p for p in documentos}
    opcoes = list(mapa)
    padrao = opcoes if selecionados_padrao is None else [p for p in opcoes if mapa[p] in selecionados_padrao]
    selecionados = st.multiselect("Documentos gerados", opcoes, default=padrao,
                                  format_func=lambda p: mapa[p].name, key=key + "_anexos") if opcoes else []
    uploads = st.file_uploader("Adicionar documentos (PDF/DOCX)", type=["pdf", "docx"], accept_multiple_files=True, key=key + "_uploads") or []
    caminhos = [mapa[p] for p in selecionados if p in mapa]
    st.caption(f"{len(caminhos)} documento(s) gerado(s) + {len(uploads)} upload(s) selecionado(s).")
    contexto = [nome_condominio, canal, destinatario, assunto, cc, bcc, telefone, mensagem]
    assinatura = _identidade_selecao(contexto, caminhos, uploads)
    estado_key = key + "_resultado"
    resultado = st.session_state.get(estado_key)
    if resultado and resultado["assinatura"] != assinatura:
        st.session_state.pop(estado_key, None)
        resultado = None
    rotulo = {"E-mail": "Enviar por e-mail", "WhatsApp": "Preparar para WhatsApp", "Ambos": "Enviar e-mail e preparar WhatsApp"}[canal]
    if resultado is None and st.button(rotulo, type="primary", use_container_width=True, key=key + "_enviar", disabled=tem_email and not cfg_ok):
        try:
            arquivos = preparar_arquivos(caminhos, uploads)
            resultado = executar_envio(canal, arquivos, enviar_email, destinatario=destinatario.strip(), assunto=assunto.strip(), mensagem=mensagem, cc=cc.strip(), bcc=bcc.strip())
            resultado.update(assinatura=assinatura, arquivos=arquivos)
            st.session_state[estado_key] = resultado
        except (ValueError, OSError) as erro:
            st.error(str(erro))
    if resultado:
        if resultado["email"] is not None:
            ok, detalhe = resultado["email"]
            (st.success if ok else st.error)("E-mail: " + detalhe)
        if resultado["whatsapp_preparado"]:
            exibir_compartilhamento(st, resultado["arquivos"], mensagem=mensagem, titulo=f"Documentos ACQUA GESTÃO - {nome_condominio}", telefone=telefone, key=key + assinatura[:12])
        if st.button("Preparar outro envio", key=key + "_novo"):
            st.session_state.pop(estado_key, None)
            st.rerun()
