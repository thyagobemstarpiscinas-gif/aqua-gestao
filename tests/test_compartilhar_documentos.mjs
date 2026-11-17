import {readFile} from 'node:fs/promises';
import assert from 'node:assert/strict';
import test from 'node:test';

const source = await readFile(new URL('../utils/compartilhar_documentos.js', import.meta.url), 'utf8');
const {default: render} = await import('data:text/javascript;base64,'+Buffer.from(source).toString('base64'));

function setup(navigator) {
  Object.defineProperty(globalThis, 'navigator', {value:navigator,configurable:true});
  const button={disabled:false}, status={textContent:''};
  const cleanup=render({parentElement:{querySelector:s=>s==='button'?button:status},
    data:{titulo:'Teste',mensagem:'<texto literal>',arquivos:[{nome:'Teste.pdf',mime:'application/pdf',base64:Buffer.from('%PDF-test').toString('base64')}]}});
  return {button,status,cleanup};
}

test('navegador sem suporte mostra alternativa sem disparar envio',()=>{
  const {button,status}=setup({});
  assert.equal(button.disabled,true);
  assert.match(status.textContent,/download/);
});
test('clique entrega bytes exatos ao compartilhamento sem anunciar entrega',async()=>{
  let payload, calls=0;
  const {button,status,cleanup}=setup({canShare:({files})=>files.length===1,share:async data=>{payload=data;calls++;}});
  assert.equal(calls,0);
  await button.onclick();
  assert.equal(calls,1);
  assert.equal(payload.files[0].name,'Teste.pdf');
  assert.equal(await payload.files[0].text(),'%PDF-test');
  assert.equal(payload.text,'<texto literal>');
  assert.match(status.textContent,/Confira a conclusão/);
  assert.equal(button.disabled,false);
  cleanup();assert.equal(button.onclick,null);
});
test('cancelamento e erro não anunciam sucesso e permitem tentar compartilhar novamente',async()=>{
  for(const name of ['AbortError','NotAllowedError']){
    let calls=0;
    const {button,status}=setup({canShare:()=>true,share:async()=>{calls++;throw Object.assign(new Error(),{name});}});
    await button.onclick();
    assert.equal(calls,1);
    assert.equal(button.disabled,false);
    assert.match(status.textContent,name==='AbortError'?/cancelado/:/Não foi possível/);
  }
});
