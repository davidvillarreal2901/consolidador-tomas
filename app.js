import {readExcel, compare, consolidate, makeOutput, name, safeText, dateKey} from './engine.js?v=20260923-2';

const $=id=>document.getElementById(id);
let state, result;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const setError=message=>{ $('error').textContent=message; $('error').hidden=!message; };
for (const [id,label] of [['previous','prev-name'],['current','curr-name']]) {
  $(id).addEventListener('change',()=>{ $(label).textContent=$(id).files[0]?.name||'Seleccionar Excel';$('results').hidden=true;$('output').hidden=true;setError(''); });
}

$('process').addEventListener('click',async()=>{
  const a=$('previous').files[0],b=$('current').files[0];
  if (!a||!b) {setError('Selecciona ambos archivos antes de comparar.');return;}
  setError('');$('process').disabled=true;$('process').textContent='Leyendo archivos…';
  try {
    const [previous,current]=await Promise.all([readExcel(a),readExcel(b)]);
    if (!previous.rows.length||!current.rows.length) throw Error('Uno de los archivos no contiene registros en las filas 16–35.');
    state=compare(previous,current); result=null; $('output').hidden=true;
    renderReview();$('results').hidden=false;$('results').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){setError('No se pudieron comparar los archivos. '+e.message);console.error(e);}
  finally{$('process').disabled=false;$('process').innerHTML='Comparar archivos <span aria-hidden="true">→</span>';}
});

function renderReview(){
  const auto=state.automatic.size, needs=state.reviews.length;
  $('stats').innerHTML=`<div class="stat"><b>${state.previous.rows.length}</b><span>En el anterior</span></div><div class="stat"><b>${state.current.rows.length}</b><span>En el actual</span></div><div class="stat"><b>${auto}</b><span>Coincidencias seguras</span></div><div class="stat"><b>${needs}</b><span>Por revisar</span></div>`;
  const found=state.current.rows.filter(r=>Array.from({length:12},(_,n)=>safeText(r.values[20+n])).some(Boolean)).length;
  const messages=[];
  if(found)messages.push(`${found} fila(s) del archivo actual ya tienen datos en Toma 2. Esos datos se conservarán como Toma 2 si la persona continúa.`);
  messages.push('La ausencia en el archivo actual se marcará como EGRESO; comprueba que hayas subido todas las páginas.');
  $('warnings').innerHTML=messages.map(esc).join('<br>');$('warnings').hidden=false;
  if(!needs){$('reviews').innerHTML='<p>No hay coincidencias dudosas. Revisa el consolidado.</p>';return;}
  $('reviews').innerHTML=`<h3>${needs} caso(s) para decidir</h3>`+state.reviews.map(({index,ranked,reason})=>{
    const row=state.current.rows[index];
    const suggestions=new Set(ranked.map(r=>r.j));
    const options=state.previous.rows.map((old,j)=>({old,j,priority:suggestions.has(j)?0:1})).sort((a,b)=>a.priority-b.priority||name(a.old).localeCompare(name(b.old),'es'));
    return `<div class="review"><div><strong>${esc(name(row)||'(Nombre incompleto)')}</strong><small>Actual · ${esc(row.id)} · NUIP ${esc(row.values[2]||'sin dato')} · nacimiento ${esc(dateKey(row.values[6])||'sin dato')}</small></div><div><label for="choice-${index}"><small>¿Corresponde a alguien anterior?</small></label><select data-review="${index}" id="choice-${index}"><option value="">Seleccionar una decisión…</option><option value="new">Nuevo ingreso</option>${options.map(({old,j})=>`<option value="${j}">${suggestions.has(j)?'★ Posible · ':''}${esc(name(old)||'(Nombre incompleto)')} · ${esc(old.values[2]||'sin NUIP')} · ${esc(dateKey(old.values[6])||'sin fecha')} · ${esc(old.id)}</option>`).join('')}</select></div><div class="reason">${esc(reason)}</div></div>`;
  }).join('');
  document.querySelectorAll('[data-review]').forEach(select=>select.addEventListener('change',()=>{
    state.assigned.delete(Number(select.dataset.review));
    if(select.value!==''&&select.value!=='new')state.assigned.set(Number(select.dataset.review),Number(select.value));
    $('output').hidden=true;
  }));
}

$('preview-button').addEventListener('click',()=>{
  if(!state)return;
  const missing=state.reviews.filter(({index})=>document.querySelector(`[data-review="${index}"]`).value==='');
  if(missing.length){setError(`Faltan ${missing.length} decisiones de la sección «Revisar coincidencias».`);$('results').scrollIntoView({behavior:'smooth'});return;}
  try{
    setError('');result=consolidate(state);
    const counts={CONTINÚA:0,NUEVO:0,EGRESO:0};for(const item of result)counts[item.status]++;
    $('summary').textContent=`${result.length} registros · ${counts['CONTINÚA']} continúan · ${counts.NUEVO} ingresos · ${counts.EGRESO} egresos · ${Math.max(1,Math.ceil(result.length/20))} hoja(s) de captura.`;
    $('preview').innerHTML=result.map(r=>`<tr><td>${r.status}</td><td>${esc(name(r.person))}</td><td>${esc(r.person.values[2])}</td><td>${esc(dateKey(r.first.values[8])||'—')}</td><td>${r.status==='EGRESO'?'EGRESO':r.status==='NUEVO'?'—':esc(dateKey(r.second.values[r.secondStart])||'—')}</td></tr>`).join('');
    $('output').hidden=false;$('output').scrollIntoView({behavior:'smooth'});
  }catch(e){setError(e.message);$('results').scrollIntoView({behavior:'smooth'});}
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

// Used for local verification with synthetic and attached test inputs.
window.__consolidador={readExcel,compare,consolidate,makeOutput};
