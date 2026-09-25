'use strict';
const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const icons = {
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5"/>',
  users: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/><circle cx="9" cy="7" r="4"/>',
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>',
  columns: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M15 3v18M3 8h18"/>',
  link: '<path d="M10 13a5 5 0 0 0 7 .1l3-3a5 5 0 0 0-7-7l-2 2M14 11a5 5 0 0 0-7-.1l-3 3a5 5 0 0 0 7 7l2-2"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  search: '<circle cx="10.5" cy="10.5" r="7.5"/><path d="m16 16 5 5"/>',
  check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
  spark: '<path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5z"/>',
};
function renderIcons(root = document) { $$('[data-icon]', root).forEach(el => { el.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[el.dataset.icon] || icons.file}</svg>`; }); }
const state = {clients: [], issuer: null, templates: [], items: [], page: 1, pages: 1, total: 0, direction: '', selected: new Set(), files: [], columns: [], config: null, reviewId: null, view: 'notas', requestId: 0};
const labels = {notas:'Notas de serviço', clientes:'Clientes', importacoes:'Importações', modelos:'Modelos de Excel', emissao:'Emitir NFS-e', integracoes:'Integrações'};
const brl = value => value == null ? 'Não informado' : Number(value).toLocaleString('pt-BR', {style:'currency',currency:'BRL'});
const dateLabel = value => value ? value.slice(0,10).split('-').reverse().join('/') : 'Não informada';
const docLabel = value => value?.length === 14 ? value.replace(/^(.{2})(.{3})(.{3})(.{4})(.{2})$/, '$1.$2.$3/$4-$5') : value || 'Não informado';
function badge(value) {
  const cls = ({Emitida:'issued', Recebida:'received', Pendente:'pending', Conferida:'done', Cancelada:'error', 'Não verificada':'neutral'})[value] || 'neutral';
  return `<span class="badge ${cls}">${escapeHtml(value)}</span>`;
}
let toastTimer;
function toast(message, error=false) { const el=$('#toast'); el.textContent=message; el.classList.toggle('error',error); el.hidden=false; clearTimeout(toastTimer); toastTimer=setTimeout(()=>el.hidden=true,5500); }
async function api(url, options={}) {
  const headers = {'X-NFSe-App':'local', ...options.headers};
  if (options.body && !(options.body instanceof FormData)) { headers['Content-Type']='application/json'; options.body=JSON.stringify(options.body); }
  let response;
  try { response = await fetch(url, {...options, headers}); } catch { throw new Error('Não foi possível conectar. Confira se a aplicação está aberta neste computador.'); }
  if (!response.ok) { const data=await response.json().catch(()=>({})); throw new Error(data.error || `Não foi possível concluir a operação (${response.status}).`); }
  return response;
}
async function json(url, options) { return (await api(url,options)).json(); }
function showDialog(id) { const dialog=$(id); $('.form-error',dialog)?.replaceChildren(); dialog.showModal(); }
function errorIn(dialog,error) { $('.form-error',$(dialog)).textContent=error.message; }
function filters() { const values=Object.fromEntries(new FormData($('#filters'))); values.direction=state.direction; return values; }
function syncDirection() { $$('.tab').forEach(button=>button.classList.toggle('active',button.dataset.direction===state.direction)); }
function updateSelection() {
  $('#selection-bar').hidden=!state.selected.size;
  $('#selection-count').textContent=`${state.selected.size} nota(s) selecionada(s)`;
  const selectedOnPage=state.items.filter(i=>state.selected.has(i.id)).length;
  $('#select-all').checked=state.items.length>0 && selectedOnPage===state.items.length;
  $('#select-all').indeterminate=selectedOnPage>0 && selectedOnPage<state.items.length;
}
async function loadClients() {
  state.clients=await json('/api/clients');
  for (const id of ['#filter-client','#import-client']) {
    const el=$(id), selected=el.value;
    el.innerHTML=(id==='#filter-client'?'<option value="">Todos os clientes</option>':'<option value="">Selecione um cliente</option>')+state.clients.map(c=>`<option value="${c.id}">${escapeHtml(c.name)}</option>`).join('');
    if (state.clients.some(c=>String(c.id)===selected)) el.value=selected;
  }
  $('#nav-count').textContent=state.clients.reduce((sum,c)=>sum+c.invoice_count,0);
  renderClients();
}
async function loadIssuer() {
  state.issuer=await json('/api/issuer');
  const form=$('#issuer-form');
  for (const name of ['name','document','city','city_code','municipal_registration','tax_regime']) {
    form.elements[name].value=state.issuer?.[name] || '';
  }
  $('#issuer-status').textContent=state.issuer ? `Configurado para ${state.issuer.name}` : 'Nenhum emitente configurado';
  $('#emission-local-guidance').innerHTML=state.issuer?.city_code==='3159605'
    ? 'Em Santa Rita do Sapucaí, a <a href="https://santaritadosapucai.mg.issqn.quasar.srv.br/issqn/" target="_blank" rel="noopener noreferrer"><u>orientação do sistema municipal</u></a> é emitir novas notas somente pelo Portal Nacional.'
    : 'Confirme no município do emitente se a emissão ocorre pelo Portal Nacional.';
}
function renderClients() {
  $('#client-list').innerHTML=state.clients.length ? state.clients.map(c=>`<article class="client-card"><span class="client-initial">${escapeHtml(c.name.slice(0,2).toUpperCase())}</span><h2>${escapeHtml(c.name)}</h2><p>${escapeHtml(docLabel(c.document))}</p><p>${escapeHtml(c.city)}</p><div class="card-footer"><span>${c.invoice_count} nota(s)</span><button class="text-button" data-client-notes="${c.id}">Ver notas →</button></div></article>`).join('') : '<div class="plain-empty"><h2>Sua carteira começa com um cliente</h2><p>Use “Novo cliente” para cadastrar o nome, CNPJ e cidade da empresa.</p></div>';
}
async function loadNotes() {
  const requestId=++state.requestId;
  const params=new URLSearchParams({...filters(),page:state.page});
  const result=await json(`/api/invoices?${params}`);
  if (requestId!==state.requestId) return;
  state.items=result.items; state.total=result.stats.total; state.pages=result.pages;
  if (state.page>state.pages) {state.page=state.pages; return loadNotes();}
  $('#stat-total').textContent=result.stats.total.toLocaleString('pt-BR');
  $('#stat-issued').textContent=brl(result.stats.amounts.Emitida);
  $('#stat-received').textContent=brl(result.stats.amounts.Recebida);
  $('#stat-pending').textContent=result.stats.pending;
  $('#invoice-rows').innerHTML=state.items.map(item=>{
    const counterparty=item.direction==='Emitida'?item.recipient_name:item.provider_name;
    const counterdoc=item.direction==='Emitida'?item.recipient_document:item.provider_document;
    return `<tr><td class="checkbox-cell"><input type="checkbox" data-select="${item.id}" aria-label="Selecionar nota ${escapeHtml(item.number || 'sem número')}" ${state.selected.has(item.id)?'checked':''}></td><td><button class="note-number" data-review="${item.id}">${item.number?'Nº '+escapeHtml(item.number):'Número a conferir'}</button><small>${escapeHtml(item.client_name)}</small></td><td>${badge(item.direction)}<small>${escapeHtml(item.source)}</small></td><td><span class="primary-text">${escapeHtml(counterparty || 'Identificação pendente')}</span><small>${escapeHtml(docLabel(counterdoc))}</small></td><td>${dateLabel(item.issued_at)}<small>${escapeHtml(item.status)}</small></td><td class="amount">${brl(item.service_value)}</td><td>${badge(item.review_status)}</td><td><button class="row-arrow" data-review="${item.id}" aria-label="Conferir nota ${escapeHtml(item.number)}">↗</button></td></tr>`;
  }).join('');
  $('#notes-empty').hidden=state.items.length>0;
  const hasAny=state.clients.some(c=>c.invoice_count>0);
  $('#notes-empty h2').textContent=hasAny?'Nenhuma nota com esses filtros':'Seu próximo fechamento começa aqui';
  $('#notes-empty p').textContent=hasAny?'Ajuste a busca ou o período para encontrar outros documentos.':'Cadastre um cliente e importe suas notas de serviço. Os documentos ficam organizados, prontos para conferir e exportar.';
  $('#empty-start').textContent=hasAny?'Limpar filtros':state.clients.length?'Importar primeiras notas':'Cadastrar primeiro cliente';
  $('#table-count').textContent=state.total?`${(state.page-1)*50+1}–${Math.min(state.page*50,state.total)} de ${state.total} nota(s)`:'Nenhuma nota na seleção';
  $('#page-label').textContent=`Página ${state.page} de ${state.pages}`;
  $('#prev-page').disabled=state.page<=1; $('#next-page').disabled=state.page>=state.pages;
  $('#export-open').disabled=!state.total;
  $('#update-caption').textContent=hasAny?'Arquivo atualizado neste computador':'Pronto para começar';
  updateSelection();
}
async function loadTemplates() {
  state.templates=await json('/api/templates');
  $('#export-template').innerHTML='<option value="">Seleção personalizada</option>'+state.templates.map(t=>`<option value="${t.id}">${escapeHtml(t.name)}</option>`).join('');
  $('#template-list').innerHTML=state.templates.length?state.templates.map(t=>`<article class="client-card"><span class="client-initial" data-icon="columns"></span><h2>${escapeHtml(t.name)}</h2><p>${t.columns.length} colunas, na ordem que você escolheu.</p><p>${escapeHtml(t.columns.slice(0,4).map(c=>c.label).join(' · '))}${t.columns.length>4?'…':''}</p><div class="card-footer"><button class="text-button" data-edit-template="${t.id}">Editar modelo →</button><button class="text-button" data-delete-template="${t.id}" aria-label="Excluir modelo ${escapeHtml(t.name)}">Excluir</button></div></article>`).join(''):'<div class="plain-empty"><h2>Seu Excel, pronto para repetir</h2><p>Crie um modelo para salvar as colunas, a ordem e os cabeçalhos de cada rotina.</p></div>';
  renderIcons($('#template-list'));
}
function resultRows(results) {return results.map(r=>`<div class="result-item"><span class="badge ${r.state==='error'?'error':r.state==='duplicate'?'neutral':'done'}">${r.state==='error'?'Atenção':r.state==='duplicate'?'Duplicado':'Pronto'}</span><div><strong>${escapeHtml(r.file)}</strong><p>${escapeHtml(r.message)}</p></div></div>`).join('');}
async function loadHistory() {
  const history=await json('/api/imports');
  $('#import-history').innerHTML=history.length?history.map(h=>`<article class="panel history-card"><div class="history-head"><h2>${escapeHtml(h.client_name)}</h2><time>${dateLabel(h.created_at)} · ${escapeHtml(h.created_at.slice(11,16))}</time></div><div class="history-stats"><span class="badge done">${h.added} nota(s) adicionada(s)</span><span class="badge neutral">${h.linked} vínculo(s)</span><span class="badge neutral">${h.duplicates} duplicado(s)</span>${h.errors?`<span class="badge error">${h.errors} erro(s)</span>`:''}</div><details><summary>Ver arquivos do lote</summary>${resultRows(h.results)}</details></article>`).join(''):'<div class="plain-empty"><h2>Nenhum arquivo importado ainda</h2><p>Cada envio aparecerá aqui com o resultado de processamento de cada arquivo.</p></div>';
}
async function navigate() {
  const view=location.hash.slice(1) || 'notas'; state.view=labels[view]?view:'notas';
  $$('.view').forEach(el=>el.hidden=el.id!==`view-${state.view}`);
  $$('.nav-item').forEach(el=>{el.classList.toggle('active',el.dataset.view===state.view); if(el.dataset.view===state.view)el.setAttribute('aria-current','page');else el.removeAttribute('aria-current');});
  $('#breadcrumb-view').textContent=labels[state.view];
  document.title=`${labels[state.view]} — Notas`;
  try {
    if (state.view==='notas') await loadNotes();
    if (state.view==='clientes') await loadClients();
    if (state.view==='emissao') await loadIssuer();
    if (state.view==='modelos') await loadTemplates();
    if (state.view==='importacoes') await loadHistory();
  } catch(error) {toast(error.message,true);}
}
function openClient() {$('#client-form').reset();showDialog('#client-dialog');}
function openImport() {
  if (!state.clients.length) {toast('Cadastre o primeiro cliente para organizar os arquivos.');openClient();return;}
  state.files=[];$('#file-input').value='';$('#file-list').innerHTML='';$('#import-result').innerHTML='';$('#import-status').textContent='Originais preservados para download';
  $('#import-submit').disabled=false;
  $('#import-client').value=$('#filter-client').value || (state.clients.length===1?state.clients[0].id:'');
  showDialog('#import-dialog');
}
function chooseFiles(files) {
  state.files=[...files]; $('#import-result').innerHTML='';$('#import-submit').disabled=false;
  $('#file-list').innerHTML=state.files.map(f=>`<div><span>${escapeHtml(f.name)}</span><span>${(f.size/1024).toLocaleString('pt-BR',{maximumFractionDigits:0})} KB</span></div>`).join('');
}
function drawColumns() {
  $('#column-count').textContent=state.columns.length;
  $('#column-choices').innerHTML=Object.entries(state.config.columns).map(([key,label])=>`<label class="column-choice"><input type="checkbox" data-column="${key}" ${state.columns.some(c=>c.key===key)?'checked':''}>${escapeHtml(label)}</label>`).join('');
  $('#column-order').innerHTML=state.columns.length?state.columns.map((c,i)=>`<div class="column-row"><span class="order-number">${i+1}</span><input data-column-label="${c.key}" value="${escapeHtml(c.label)}" maxlength="100" aria-label="Cabeçalho de ${escapeHtml(state.config.columns[c.key])}"><button class="icon-button" data-move="${i}" data-step="-1" aria-label="Mover ${escapeHtml(c.label)} para cima" ${i===0?'disabled':''}>↑</button><button class="icon-button" data-move="${i}" data-step="1" aria-label="Mover ${escapeHtml(c.label)} para baixo" ${i===state.columns.length-1?'disabled':''}>↓</button></div>`).join(''):'<p class="muted">Selecione as colunas ao lado para montar sua planilha.</p>';
}
function openExport(templateOnly=false,template=null) {
  state.columns=template?structuredClone(template.columns):state.config.default_columns.map(key=>({key,label:state.config.columns[key]}));
  $('#export-title').textContent=templateOnly?'Modelo de Excel':'Preparar Excel';
  $('#export-template').value=template?.id || '';$('#template-name').value=template?.name || '';
  $('#export-scope-label').hidden=templateOnly;$('#export-download').hidden=templateOnly;$('#export-warning').hidden=templateOnly;
  $('#export-scope').value=state.selected.size?'selected':'filtered';
  $('#export-scope option[value="selected"]').disabled=!state.selected.size;
  drawColumns();showDialog('#export-dialog');
}
function reviewField(item,key,type='text',full=false) {
  const label=state.config.columns[key], val=item[key] ?? '';
  let input;
  if (type==='textarea') input=`<textarea name="${key}" maxlength="10000">${escapeHtml(val)}</textarea>`;
  else if (Array.isArray(type)) input=`<select name="${key}">${type.map(option=>`<option value="${escapeHtml(option)}" ${option===val?'selected':''}>${escapeHtml(option || 'Não informado')}</option>`).join('')}</select>`;
  else input=`<input name="${key}" type="${type==='money'?'text':type}" ${type==='money'?'inputmode="decimal"':''} value="${escapeHtml(type==='money'&&val!==''?val.replace('.',','):val)}" ${type==='date'?'':'maxlength="10000"'}>`;
  return `<label class="${full?'full':''}">${escapeHtml(label)}${input}</label>`;
}
async function openReview(id) {
  try {
    const item=await json(`/api/invoices/${id}`);state.reviewId=id;
    $('#review-title').textContent=item.number?`Nota nº ${item.number}`:'Nota para identificar';
    $('#review-summary').innerHTML=`<div class="review-meta"><strong>${escapeHtml(item.client_name)}</strong>${badge(item.direction)}${badge(item.review_status)}<span class="badge neutral">${escapeHtml(item.source)}</span></div>${item.access_key?`<p class="review-key">Chave: ${escapeHtml(item.access_key)}</p>`:''}`;
    $('#review-attachments').innerHTML=item.attachments.map(a=>`<a class="attachment-link" href="/api/attachments/${a.id}" download><span>↓ ${escapeHtml(a.kind)}</span>${escapeHtml(a.name)}</a>`).join('');
    const group=(title,fields)=>`<h3>${title}</h3><div class="review-grid">${fields.join('')}</div>`;
    $('#review-fields').innerHTML=
      group('Dados da nota', [reviewField(item,'number'),reviewField(item,'verification_code'),reviewField(item,'issued_at','date'),reviewField(item,'competence','date'),reviewField(item,'municipality'),reviewField(item,'status',['Não verificada','Normal','Cancelada','Substituída'])])+
      group('Quem prestou e quem recebeu o serviço', ['provider_name','provider_document','recipient_name','recipient_document'].map(k=>reviewField(item,k)))+
      '<p class="muted">O tipo é identificado pelo CPF/CNPJ das partes. Para vincular ao cliente, o documento deve corresponder ao CNPJ cadastrado.</p>'+
      group('Serviço e valores informados', [reviewField(item,'service_code'),reviewField(item,'iss_withheld',['','Sim','Não']),reviewField(item,'description','textarea',true),...['service_value','net_value','deductions','iss','pis','cofins','inss','ir','csll'].map(k=>reviewField(item,k,'money')),reviewField(item,'notes','textarea',true)]);
    $('#pdf-text-panel').hidden=!item.raw_text;$('#pdf-text-panel').open=false;$('#pdf-text').textContent=item.raw_text || '';
    $('#review-complete').checked=item.review_status==='Conferida';showDialog('#review-dialog');
  }catch(error){toast(error.message,true);}
}

document.addEventListener('click',event=>{
  const close=event.target.closest('.close-dialog');if(close)close.closest('dialog').close();
  const review=event.target.closest('[data-review]');if(review)openReview(Number(review.dataset.review));
  const client=event.target.closest('[data-client-notes]');if(client){$('#filters').reset();$('#filter-client').value=client.dataset.clientNotes;state.direction='';syncDirection();state.page=1;state.selected.clear();location.hash='notas';if(state.view==='notas')loadNotes().catch(e=>toast(e.message,true));}
  const edit=event.target.closest('[data-edit-template]');if(edit)openExport(true,state.templates.find(t=>t.id===Number(edit.dataset.editTemplate)));
  const del=event.target.closest('[data-delete-template]');if(del){const id=Number(del.dataset.deleteTemplate);api(`/api/templates/${id}`,{method:'DELETE'}).then(loadTemplates).then(()=>toast('Modelo excluído.')).catch(e=>toast(e.message,true));}
  const move=event.target.closest('[data-move]');if(move){const i=Number(move.dataset.move),to=i+Number(move.dataset.step);const key=state.columns[i].key;[state.columns[i],state.columns[to]]=[state.columns[to],state.columns[i]];drawColumns();$(`[data-column-label="${key}"]`).focus();$('#export-template').value='';}
});
window.addEventListener('hashchange',navigate);
for(const id of ['#import-open','#history-import'])$(id).addEventListener('click',openImport);
$('#client-open').addEventListener('click',openClient);
$('#empty-start').addEventListener('click',()=>{if(state.clients.some(c=>c.invoice_count))clearFilters();else if(state.clients.length)openImport();else openClient();});
$('#export-open').addEventListener('click',()=>openExport());
$('#template-new').addEventListener('click',()=>openExport(true));
$('#client-form').addEventListener('submit',async event=>{
  event.preventDefault();const button=$('button[type="submit"]',event.target);button.disabled=true;
  try {const result=await json('/api/clients',{method:'POST',body:Object.fromEntries(new FormData(event.target))});await loadClients();$('#client-dialog').close();toast('Cliente cadastrado. Você já pode importar as notas.');if(state.view==='notas'){await loadNotes();openImport();$('#import-client').value=result.id;}}
  catch(error){errorIn('#client-dialog',error);}finally{button.disabled=false;}
});
$('#issuer-form').addEventListener('submit',async event=>{
  event.preventDefault();
  const button=$('button[type="submit"]',event.target);button.disabled=true;
  $('.form-error',event.target).textContent='';
  try {await json('/api/issuer',{method:'PUT',body:Object.fromEntries(new FormData(event.target))});await loadIssuer();toast('Dados do emitente salvos nesta instalação.');}
  catch(error){errorIn('#issuer-form',error);}finally{button.disabled=false;}
});
$('#file-input').addEventListener('change',event=>chooseFiles(event.target.files));
for (const name of ['dragover','dragenter'])$('#dropzone').addEventListener(name,event=>{event.preventDefault();$('#dropzone').classList.add('dragging');});
$('#dropzone').addEventListener('dragleave',()=>$('#dropzone').classList.remove('dragging'));
$('#dropzone').addEventListener('drop',event=>{event.preventDefault();$('#dropzone').classList.remove('dragging');chooseFiles(event.dataTransfer.files);});
$('#import-form').addEventListener('submit',async event=>{
  event.preventDefault();$('.form-error',event.target).textContent='';
  if(!state.files.length){errorIn('#import-dialog',new Error('Selecione os arquivos para importar.'));return;}
  if(state.files.some(f=>f.size>20*1024*1024)||state.files.reduce((sum,f)=>sum+f.size,0)>59*1024*1024){errorIn('#import-dialog',new Error('Limite de 20 MB por arquivo e 60 MB por envio. Divida o lote.'));return;}
  const button=$('#import-submit');button.disabled=true;button.textContent='Processando…';$('#import-status').textContent='Lendo e organizando os documentos…';
  // Prevent closing a running import and starting a concurrent batch accidentally.
  const dialog=$('#import-dialog');dialog.dataset.busy='true';$('.close-dialog',dialog).disabled=true;$('#import-client').disabled=true;$('#file-input').disabled=true;
  try{
    const data=new FormData();data.append('client_id',$('#import-client').value);state.files.forEach(file=>data.append('files',file));
    const result=await json('/api/imports',{method:'POST',body:data});
    $('#import-result').innerHTML=`<div class="note-box"><strong>${result.added} nota(s) adicionada(s)</strong> · ${result.linked} vínculo(s) · ${result.duplicates} duplicado(s) · ${result.errors} erro(s)</div>${resultRows(result.results)}`;
    $('#import-status').textContent='Processamento concluído';state.files=[];$('#file-input').value='';$('#file-list').innerHTML='';
    await loadClients();await loadNotes();if(state.view==='importacoes')await loadHistory();
  }catch(error){errorIn('#import-dialog',error);$('#import-status').textContent='O envio não foi concluído';button.disabled=false;}
  finally{button.textContent='Importar arquivos →';dialog.dataset.busy='false';$('.close-dialog',dialog).disabled=false;$('#import-client').disabled=false;$('#file-input').disabled=false;}
});
$('#import-dialog').addEventListener('cancel',event=>{if(event.currentTarget.dataset.busy==='true')event.preventDefault();});
let searchTimer;
function changedFilters(){state.page=1;state.selected.clear();loadNotes().catch(e=>toast(e.message,true));}
function clearFilters(){ $('#filters').reset(); }
$('#filters').addEventListener('submit',event=>event.preventDefault());
$('#filters').addEventListener('input',event=>{if(event.target.name==='q'){clearTimeout(searchTimer);searchTimer=setTimeout(changedFilters,250);}});
$('#filters').addEventListener('change',event=>{if(event.target.name!=='q')changedFilters();});
$('#filters').addEventListener('reset',()=>{state.direction='';syncDirection();setTimeout(changedFilters,0);});
$$('.tab').forEach(button=>button.addEventListener('click',()=>{state.direction=button.dataset.direction;syncDirection();changedFilters();}));
$('#pending-filter').addEventListener('click',()=>{$('#filters [name="review_status"]').value='Pendente';changedFilters();});
$('#prev-page').addEventListener('click',()=>{state.page--;loadNotes().catch(e=>toast(e.message,true));});
$('#next-page').addEventListener('click',()=>{state.page++;loadNotes().catch(e=>toast(e.message,true));});
$('#invoice-rows').addEventListener('change',event=>{if(event.target.dataset.select){const id=Number(event.target.dataset.select);event.target.checked?state.selected.add(id):state.selected.delete(id);updateSelection();}});
$('#select-all').addEventListener('change',event=>{state.items.forEach(i=>event.target.checked?state.selected.add(i.id):state.selected.delete(i.id));$$('[data-select]').forEach(el=>el.checked=state.selected.has(Number(el.dataset.select)));updateSelection();});
$('#clear-selection').addEventListener('click',()=>{state.selected.clear();$$('[data-select]').forEach(el=>el.checked=false);updateSelection();});
$('#column-choices').addEventListener('change',event=>{const key=event.target.dataset.column;if(!key)return;if(event.target.checked)state.columns.push({key,label:state.config.columns[key]});else state.columns=state.columns.filter(c=>c.key!==key);drawColumns();$(`[data-column="${key}"]`).focus();$('#export-template').value='';});
$('#column-order').addEventListener('input',event=>{const column=state.columns.find(c=>c.key===event.target.dataset.columnLabel);if(column){column.label=event.target.value;$('#export-template').value='';}});
$('#export-template').addEventListener('change',event=>{const template=state.templates.find(t=>String(t.id)===event.target.value);if(template){state.columns=structuredClone(template.columns);$('#template-name').value=template.name;drawColumns();}});
$('#template-save').addEventListener('click',async event=>{
  event.target.disabled=true;$('.form-error',$('#export-dialog')).textContent='';
  try{await json('/api/templates',{method:'POST',body:{name:$('#template-name').value,columns:state.columns}});await loadTemplates();toast('Modelo salvo. Ele estará disponível nos próximos acessos.');$('#export-template').value=state.templates.find(t=>t.name===$('#template-name').value.trim())?.id || '';}
  catch(error){errorIn('#export-dialog',error);}finally{event.target.disabled=false;}
});
$('#export-download').addEventListener('click',async()=>{
  const button=$('#export-download');button.disabled=true;$('.form-error',$('#export-dialog')).textContent='';
  try{const payload={columns:state.columns,filters:filters()};if($('#export-scope').value==='selected')payload.ids=[...state.selected];const response=await api('/api/export',{method:'POST',body:payload});const blob=await response.blob();const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=`notas-servico-${new Date().toISOString().slice(0,10)}.xlsx`;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),5000);$('#export-dialog').close();toast('Planilha gerada. Confira o download no navegador.');}
  catch(error){errorIn('#export-dialog',error);}finally{button.disabled=false;}
});
$('#review-form').addEventListener('submit',async event=>{
  event.preventDefault();const button=$('button[type="submit"]',event.target);button.disabled=true;
  try{const data=Object.fromEntries(new FormData(event.target));data.review_status=$('#review-complete').checked?'Conferida':'Pendente';await json(`/api/invoices/${state.reviewId}`,{method:'PATCH',body:data});$('#review-dialog').close();toast('Conferência salva. Os arquivos originais foram preservados.');await loadNotes();}
  catch(error){errorIn('#review-dialog',error);}finally{button.disabled=false;}
});
async function init(){renderIcons();try{state.config=await json('/api/config');await Promise.all([loadClients(),loadTemplates(),loadIssuer()]);await navigate();}catch(error){toast(error.message,true);$('#notes-empty h2').textContent='Não foi possível carregar os dados';$('#notes-empty p').textContent=error.message;}}
init();
