// Dados de documentos ficam apenas nesta página; nenhum link público é criado.
export default function(component) {
  const {parentElement, data} = component;
  const button = parentElement.querySelector('button');
  const status = parentElement.querySelector('[role="status"]');
  let files = [];
  try {
    files = data.arquivos.map(item => new File(
      [Uint8Array.from(atob(item.base64), ch => ch.charCodeAt(0))],
      item.nome, {type: item.mime}
    ));
    if (!navigator.share || !navigator.canShare || !navigator.canShare({files})) {
      button.disabled = true;
      status.textContent = 'Este navegador não permite compartilhar estes arquivos diretamente. Use as opções de download abaixo ou abra o app no navegador do celular.';
    }
  } catch (_) {
    button.disabled = true;
    status.textContent = 'Não foi possível preparar o compartilhamento neste navegador. Os arquivos estão disponíveis para download abaixo.';
  }
  button.onclick = async () => {
    button.disabled = true;
    status.textContent = 'Escolha WhatsApp, selecione o contato e confirme o envio.';
    try {
      // Chamado no clique, sem fetch/await anterior que consumiria a ativação do usuário.
      await navigator.share({files, title: data.titulo, text: data.mensagem});
      status.textContent = 'Documentos encaminhados ao compartilhamento do aparelho. Confira a conclusão do envio no WhatsApp.';
    } catch (error) {
      status.textContent = error.name === 'AbortError'
        ? 'Compartilhamento cancelado ou indisponível. O app não confirmou envio pelo WhatsApp.'
        : 'Não foi possível compartilhar. Use o download abaixo para anexar o arquivo no WhatsApp.';
    } finally {
      button.disabled = false;
    }
  };
  return () => { button.onclick = null; files = []; };
}
