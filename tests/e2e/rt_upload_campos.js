// Navegador real: relatório RT do escritório. Para cada cliente fictício: carrega o cliente,
// envia fotos (1, depois 2), exclui uma, provoca reruns e confere que TODOS os campos de texto
// da página continuam com os mesmos valores.
// Alfa: visita fora do mês (cenário do defeito). Beta: visita no mês (importação automática).
// Uso: sh tests/e2e/rodar_rt_upload.sh  (app isolado, planilha falsa, sem rede, dados fictícios)
const { chromium } = require('playwright');
const [PORTA, PIN] = [process.argv[2], process.argv[3]];
// PNG 1x1 fictício (nenhuma foto real)
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==', 'base64');
const foto = (n) => ({ name: `foto_${n}.png`, mimeType: 'image/png', buffer: PNG });
const PRINCIPAIS = ['Condomínio / estabelecimento', 'Endereço do local', 'Representante / síndico / contato local',
  'Mês de referência', 'Ano de referência'];

async function cenario(b, cliente, esperaAnalise) {
  const p = await b.newPage({ viewport: { width: 1400, height: 3200 } });
  const espera = async () => { await p.waitForTimeout(1000); await p.waitForFunction(() => !document.querySelector('[data-testid="stStatusWidget"]'), null, { timeout: 120000 }).catch(() => {}); await p.waitForTimeout(1200); };
  // Todos os campos de texto da página (ocorrência + rótulo), exceto o PIN.
  const campos = async () => Object.fromEntries((await p.$$eval('input[type=text], textarea', els => {
    const n = {}; return els.map(e => { const r = e.getAttribute('aria-label') || ''; n[r] = (n[r] || 0) + 1; return [`${n[r]}:${r}`, e.value]; });
  })).filter(([k]) => !/PIN/i.test(k)));
  let inicial;
  const up = () => p.locator('[data-testid="stFileUploader"]').filter({ hasText: 'Upload de fotos do relatório' });
  const nFotos = () => up().locator('[data-testid="stFileChip"]').count();
  let falhas = 0; let ref;
  const conferir = async (rot, permitido) => {  // permitido: rótulo do campo que a etapa altera de propósito
    const v = await campos(); const dif = Object.keys(ref).filter(k => v[k] !== ref[k] && !(permitido && k.endsWith(':' + permitido)));
    for (const c of PRINCIPAIS) if (v['1:' + c] !== inicial[c] && !dif.includes('1:' + c)) dif.push('1:' + c);
    if (dif.length) falhas++;
    console.log(`  ${dif.length ? 'FALHOU' : 'OK    '} ${rot.padEnd(30)} fotos=${await nFotos()} campos=${Object.keys(v).length}`
      + (dif.length ? ` alterados(${dif.length}): ` + dif.slice(0, 6).map(k => `${k.replace(/^\d+:/, '')}=${JSON.stringify(v[k])}`).join('; ') : ''));
    ref = v;
  };
  console.log(`== ${cliente}`);
  await p.goto(`http://127.0.0.1:${PORTA}/`); await p.getByRole('button', { name: '🔐 Acesso administrativo' }).click(); await espera();
  await p.getByRole('button', { name: '🔵 Aqua Gestão' }).click(); await espera();
  await p.getByPlaceholder('Digite o PIN administrativo').fill(PIN); await p.getByRole('button', { name: 'Entrar no escritório' }).click(); await espera(); await espera();
  const box = p.locator('[data-testid="stHorizontalBlock"]').filter({ has: p.getByRole('button', { name: '⬇️ Carregar', exact: true }) }).last().locator('[data-testid="stSelectbox"]').first();
  await box.scrollIntoViewIfNeeded(); await box.getByRole('button', { name: 'Open' }).click(); await p.waitForTimeout(800);
  await p.getByRole('option').filter({ hasText: cliente }).first().click(); await espera();
  await p.getByRole('button', { name: '⬇️ Carregar', exact: true }).click(); await espera(); await espera();
  ref = await campos();
  const principais = {}; for (const c of PRINCIPAIS) principais[c] = await p.getByLabel(c, { exact: true }).inputValue();
  console.log('  após Carregar:', JSON.stringify(principais)); inicial = principais;
  if (PRINCIPAIS.some(c => !principais[c])) { falhas++; console.log('  FALHOU campo principal vazio na tela logo após Carregar'); }
  if (esperaAnalise && !Object.values(ref).includes(esperaAnalise)) { falhas++; console.log(`  FALHOU análise ${esperaAnalise} não importada`); }
  const inp = up().locator('input[type=file]');
  await inp.setInputFiles(foto(0)); await espera(); await espera(); await conferir('1 foto');
  await inp.setInputFiles([foto(1), foto(2)]); await espera(); await espera(); await conferir('+2 fotos');
  await up().locator('[data-testid="stFileChipDeleteBtn"] button, [data-testid="stFileChipDeleteBtn"]').first().click({ timeout: 10000 }); await espera(); await espera(); await conferir('excluir 1 foto');
  await p.locator('input[aria-label="Data de emissão"][placeholder="dd/mm/aaaa"]').last().fill('26/10/2026'); await p.keyboard.press('Tab'); await espera();
  await conferir('rerun: editar data de emissão', 'Data de emissão');
  await p.getByRole('checkbox', { name: 'Salvar alterações deste relatório no cadastro principal' }).click({ force: true }).catch(() => {}); await espera(); await conferir('rerun: checkbox');
  await inp.setInputFiles(foto(3)); await espera(); await espera(); await conferir('mais 1 foto após reruns');
  if (await nFotos() !== 3) { falhas++; console.log('  FALHOU fotos anexadas diferente de 3'); }
  await p.close();
  return falhas;
}

(async () => {
  const b = await chromium.launch();
  let falhas = await cenario(b, 'Alfa', null);
  falhas += await cenario(b, 'Beta', '7,2');
  console.log(falhas ? `RESULTADO: ${falhas} FALHA(S)` : 'RESULTADO: todos os campos preservados');
  await b.close(); process.exit(falhas ? 1 : 0);
})();
