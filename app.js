import {readExcel, compare, consolidate, makeOutput, outputValue, sortConsolidated, pagePlan, name, safeText, dateKey} from './engine.js?v=20260923-4';

const $=id=>document.getElementById(id);
let state, result, editingIndex=null;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const loc=r=>`hoja «${r.sheet.title}», fila ${r.row}`;
const cell=(c,row)=>{let letters='';while(c){c--;letters=String.fromCharCode(65+c%26)+letters;c=Math.floor(c/26);}return `${letters}${row}`;};
const setError=message=>{$('error').textContent=message;$('error').hidden=!message;};
const displayDate=v=>{const d=dateKey(v);return /^\d{4}-\d\d-\d\d$/.test(d)?`${d.slice(8,10)}/${d.slice(5,7)}/${d.slice(0,4)}`:safeText(v);};
const validDate=v=>{const d=dateKey(v),m=d.match(/^(\d{4})-(\d\d)-(\d\d)$/);return Boolean(m&&new Date(Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3]))).toISOString().slice(0,10)===d);};
const columns=['Fecha de la toma','Peso (kg)','Talla (cm)','Perímetro braquial (cm)','Alimentación en menores de 6 meses','Lactancia exclusiva (meses)','Recibe leche humana','Edad de lactancia exclusiva','Edad al introducir alimentos','Lactancia total (meses)','Textura de primeros alimentos','Utensilios de alimentación'];
const groups=[{name:'Identificación',start:2,labels:['NUIP','Nombres','Apellidos','Sexo','Fecha de nacimiento','Fecha de ingreso']},{name:'Toma 1',start:8,labels:columns},{name:'Toma 2',start:20,labels:columns}];

for(const [id,label] of [['previous','prev-name'],['current','curr-name']]){
  $(id).addEventListener('change',()=>{ $(label).textContent=$(id).files[0]?.name||'Seleccionar Excel';state=undefined;result=null;$('results').hidden=true;$('output').hidden=true;setError(''); });
}

$('process').addEventListener('click',async()=>{
  const a=$('previous').files[0],b=$('current').files[0];
  if(!a||!b){setError('Selecciona ambos archivos antes de comparar.');return;}
  setError('');$('process').disabled=true;$('process').textContent='Leyendo archivos…';
  try{
    const [previous,current]=await Promise.all([readExcel(a),readExcel(b)]);
    if(!previous.rows.length||!current.rows.length)throw Error('Uno de los archivos no contiene filas de niños. Revisa las hojas y la columna de número de orden.');
    state=compare(previous,current);result=null;$('output').hidden=true;
    renderReview();$('results').hidden=false;$('results').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){setError('No se pudieron comparar los archivos. '+e.message);console.error(e);}
  finally{$('process').disabled=false;$('process').innerHTML='Comparar archivos <span aria-hidden="true">→</span>';}
});

function renderReview(){
  const auto=state.automatic.size,needs=state.reviews.length;
  $('stats').innerHTML=`<div class="stat"><b>${state.previous.rows.length}</b><span>En el anterior</span></div><div class="stat"><b>${state.current.rows.length}</b><span>En el actual</span></div><div class="stat"><b>${auto}</b><span>Coincidencias seguras</span></div><div class="stat"><b>${needs}</b><span>Por revisar</span></div>`;
  const messages=['Se ignora por completo la Toma 2 del archivo anterior. Solo se usa para Toma 2 la medición del archivo actual cuando corresponde.','La ausencia en el archivo actual se marcará como EGRESO; comprueba que hayas subido todas las páginas.'];
  $('warnings').innerHTML=messages.map(esc).join('<br>');$('warnings').hidden=false;
  if(!needs){$('reviews').innerHTML='<p>No hay coincidencias dudosas. Revisa el consolidado.</p>';return;}
  $('reviews').innerHTML=`<h3>${needs} caso(s) para decidir</h3>`+state.reviews.map(({index,ranked,reason})=>{
    const row=state.current.rows[index],best=ranked.length?state.previous.rows[ranked[0].j]:null;
    const suggestions=new Set(ranked.map(r=>r.j));
    const options=state.previous.rows.map((old,j)=>({old,j,priority:suggestions.has(j)?0:1})).sort((a,b)=>a.priority-b.priority||name(a.old).localeCompare(name(b.old),'es'));
    const origins=`Actual: ${loc(row)} · NUIP ${cell(2,row.row)} · Nombres ${cell(3,row.row)} · Apellidos ${cell(4,row.row)} · Nacimiento ${cell(6,row.row)}`;
    const prior=best?`Posible anterior: ${loc(best)} · NUIP ${best.values[2]||'sin dato'} (${cell(2,best.row)}) · ${name(best)||'nombre incompleto'} (${cell(3,best.row)} y ${cell(4,best.row)}) · nacimiento ${dateKey(best.values[6])||'sin dato'} (${cell(6,best.row)})`:'';
    return `<div class="review"><div><strong>${esc(name(row)||'(Nombre incompleto)')}</strong><small>Actual · ${esc(loc(row))} · NUIP ${esc(row.values[2]||'sin dato')} · nacimiento ${esc(dateKey(row.values[6])||'sin dato')}</small></div><div><label for="choice-${index}"><small>¿Corresponde a alguien anterior?</small></label><select data-review="${index}" id="choice-${index}"><option value="">Seleccionar una decisión…</option><option value="new">Nuevo ingreso</option>${options.map(({old,j})=>`<option value="${j}">${suggestions.has(j)?'★ Posible · ':''}${esc(name(old)||'(Nombre incompleto)')} · ${esc(old.values[2]||'sin NUIP')} · ${esc(loc(old))}</option>`).join('')}</select></div><div class="reason">${esc(reason)}<div class="details">${esc(origins)}${prior?'<br>'+esc(prior):''}</div></div></div>`;
  }).join('');
  document.querySelectorAll('[data-review]').forEach(select=>select.addEventListener('change',()=>{
    state.assigned.delete(Number(select.dataset.review));
    if(select.value!==''&&select.value!=='new')state.assigned.set(Number(select.dataset.review),Number(select.value));
    result=null;$('output').hidden=true;
  }));
}

function renderPreview(){
  sortConsolidated(result);
  const counts={CONTINÚA:0,NUEVO:0,EGRESO:0};for(const item of result)counts[item.status]++;
  const corrections=result.filter(item=>Object.keys(item.overrides||{}).length).length;
  $('summary').textContent=`${result.length} registros · ${counts['CONTINÚA']} continúan · ${counts.NUEVO} ingresos · ${counts.EGRESO} egresos · ${pagePlan(state.current,result.length).length} hoja(s) · ${corrections} registro(s) corregido(s).`;
  $('preview').innerHTML=result.map((item,i)=>{
    const given=safeText(outputValue(item,3)),surname=safeText(outputValue(item,4));
    const second=item.status==='EGRESO'?'EGRESO':item.status==='NUEVO'?'—':dateKey(outputValue(item,20))||'—';
    const source=item.status==='CONTINÚA'?`Actual: ${loc(item.person)}; anterior: ${loc(item.first)}`:`${item.status==='EGRESO'?'Anterior':'Actual'}: ${loc(item.person)}`;
    return `<tr><td>${item.status}</td><td>${esc(`${given} ${surname}`.trim())}${item.overrides&&Object.keys(item.overrides).length?'<span class="edited">✓ Corregido</span>':''}</td><td>${esc(outputValue(item,2))}</td><td>${esc(dateKey(outputValue(item,8))||'—')}</td><td>${esc(second)}</td><td class="source-cell">${esc(source)}</td><td><button type="button" class="edit-link" data-edit="${i}">Editar</button></td></tr>`;
  }).join('');
}

$('preview-button').addEventListener('click',()=>{
  if(!state)return;
  const missing=state.reviews.filter(({index})=>document.querySelector(`[data-review="${index}"]`).value==='');
  if(missing.length){setError(`Faltan ${missing.length} decisiones de la sección «Revisar coincidencias».`);$('results').scrollIntoView({behavior:'smooth'});return;}
  try{
    setError('');if(!result)result=consolidate(state);
    renderPreview();$('output').hidden=false;$('output').scrollIntoView({behavior:'smooth'});
  }catch(e){setError(e.message);$('results').scrollIntoView({behavior:'smooth'});}
});

function fieldSource(item,c){
  if(c<=7)return {row:item.person,col:c};
  if(c<=19)return {row:item.first,col:c};
  if(!item.second)return null;
  return {row:item.second,col:item.secondStart+c-20};
}
function inputValue(item,c){const v=outputValue(item,c);return [6,7,8,20].includes(c)?displayDate(v):safeText(v);}

$('preview').addEventListener('click',event=>{
  const button=event.target.closest('[data-edit]');if(!button||!result)return;
  editingIndex=Number(button.dataset.edit);
  const item=result[editingIndex];
  $('edit-title').textContent=`Editar ${safeText(outputValue(item,3))} ${safeText(outputValue(item,4))}`.trim();
  $('edit-origin').textContent=`${item.status} · ${loc(item.person)}. Estos cambios solo afectan al Excel que descargues.`;
  $('edit-error').hidden=true;
  $('edit-fields').innerHTML=groups.filter(g=>g.start!==20||item.status!=='EGRESO').map(g=>`<fieldset class="edit-group"><legend>${g.name}</legend><div class="edit-grid">${g.labels.map((label,k)=>{
    const c=g.start+k,source=fieldSource(item,c),origin=source?`${source.row.sheet.title}!${cell(source.col,source.row.row)}`:'Sin dato de origen';
    return `<label class="edit-field">${esc(label)} <small>${esc(origin)}</small><input type="text" data-column="${c}" value="${esc(inputValue(item,c))}" spellcheck="false" autocomplete="off"></label>`;
  }).join('')}</div></fieldset>`).join('');
  $('edit-dialog').showModal();
});

for(const id of ['edit-cancel','edit-dismiss'])$(id).addEventListener('click',()=>$('edit-dialog').close());
$('edit-form').addEventListener('submit',event=>{
  event.preventDefault();
  const item=result[editingIndex],updates={};
  for(const input of $('edit-fields').querySelectorAll('[data-column]')){
    const c=Number(input.dataset.column),value=input.value.trim();
    if([6,8,20].includes(c)&&value&&value!==inputValue(item,c)&&!validDate(value)){
      $('edit-error').textContent=`Revisa la fecha de «${input.closest('label').firstChild.textContent.trim()}». Usa día/mes/año.`;
      $('edit-error').hidden=false;input.focus();return;
    }
    if(value!==inputValue(item,c))updates[c]=value;
  }
  item.overrides={...(item.overrides||{}),...updates};
  $('edit-dialog').close();renderPreview();
});

$('download').addEventListener('click',async()=>{
  if(!result)return;
  const button=$('download');button.disabled=true;button.textContent='Preparando Excel…';
  try{
    const blob=await makeOutput(state,result),url=URL.createObjectURL(blob),link=document.createElement('a');
    link.href=url;link.download='Consolidado_tomas.xlsx';document.body.append(link);link.click();link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),60000);
  }catch(e){setError('No se pudo generar el Excel. '+e.message);console.error(e);}
  finally{button.disabled=false;button.textContent='Descargar Excel consolidado ↓';}
});

window.__consolidador={readExcel,compare,consolidate,makeOutput};
