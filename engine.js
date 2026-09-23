// Excel Open XML reader/writer. JSZip is vendored; all processing stays in the browser.
const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
const C = 'http://schemas.openxmlformats.org/package/2006/content-types';
const parser = new DOMParser();
const serializer = new XMLSerializer();
const children = (node, tag) => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1 && (!tag || n.localName === tag));
const first = (node, tag) => children(node, tag)[0];
const xml = text => {
  const doc = parser.parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw Error('El archivo contiene XML inválido.');
  return doc;
};
const col = letters => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const letter = num => { let s = ''; while (num) { num--; s = String.fromCharCode(65 + num % 26) + s; num = Math.floor(num / 26); } return s; };
const addr = (c, r) => `${letter(c)}${r}`;
const safeText = x => String(x ?? '').trim();

function cellValue(cell, strings) {
  if (!cell) return '';
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') return Array.from(cell.getElementsByTagNameNS(M, 't')).map(n => n.textContent).join('');
  const value = first(cell, 'v')?.textContent ?? '';
  if (type === 's') return strings[Number(value)] ?? '';
  if (type === 'b') return value === '1' ? 'SI' : 'NO';
  return value;
}

function dateKey(v) {
  v = safeText(v);
  if (!v) return '';
  if (/^\d{4}-\d\d-\d\d/.test(v)) return v.slice(0, 10);
  let m = v.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  if (/^\d+(?:\.\d+)?$/.test(v)) {
    const ms = (Number(v) - 25569) * 86400000;
    if (ms > -2208988800000 && ms < 4102444800000) return new Date(ms).toISOString().slice(0, 10);
  }
  return v.toUpperCase();
}
const key = x => safeText(x).replace(/\s+/g, ' ').toLocaleUpperCase('es-CO');
const docKey = x => safeText(x).replace(/[\s.\-]/g, '').toUpperCase();
const name = x => `${safeText(x.values[3])} ${safeText(x.values[4])}`.trim();
const identity = x => `${key(x.values[3])}|${key(x.values[4])}|${dateKey(x.values[6])}`;
const firstName = x => key(x.values[3]).split(' ')[0];
const isInstructions = title => /^instruc+iones$/i.test(key(title).normalize('NFD').replace(/[\u0300-\u036f\s]/g,''));
const isOrder = v => /^\d{1,5}$/.test(safeText(v)) && Number(v)>=1;

function sheetLayout(doc, strings) {
  const rows=children(first(doc.documentElement,'sheetData'),'row');
  const at=(row,c)=>cellValue(children(row,'c').find(x=>x.getAttribute('r')===addr(c,Number(row.getAttribute('r')))),strings);
  const header=rows.find(row=>/NO\.?\s*DE\s*ORDEN/i.test(safeText(at(row,1))))?.getAttribute('r');
  const headerRow=Number(header)||14;
  const firstNumber=rows.find(row=>Number(row.getAttribute('r'))>headerRow && Number(row.getAttribute('r'))<=headerRow+5 && isOrder(at(row,1)));
  const start=firstNumber?Number(firstNumber.getAttribute('r')):headerRow+2;
  const note=rows.find(row=>Number(row.getAttribute('r'))>=start && Number(row.getAttribute('r'))<=start+20 && /^NOTAS?\s*:/i.test(key(at(row,1))));
  const capacity=Math.min(20,(note?Number(note.getAttribute('r')):start+20)-start);
  if(capacity<1)throw Error('La hoja no tiene espacio para registros de niños.');
  return {headerRow,start,capacity};
}

async function loadXml(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw Error(`Falta ${path} en el archivo Excel.`);
  return xml(await entry.async('string'));
}

export async function readExcel(file) {
  const zip = await JSZip.loadAsync(file);
  const book = await loadXml(zip, 'xl/workbook.xml');
  const rels = await loadXml(zip, 'xl/_rels/workbook.xml.rels');
  const map = new Map(children(rels.documentElement, 'Relationship').map(n => [n.getAttribute('Id'), n.getAttribute('Target')]));
  const ssFile = zip.file('xl/sharedStrings.xml');
  const strings = ssFile ? children((await loadXml(zip, 'xl/sharedStrings.xml')).documentElement, 'si').map(si => Array.from(si.getElementsByTagNameNS(M, 't')).map(t => t.textContent).join('')) : [];
  const sheets = [];
  const styles = await loadXml(zip,'xl/styles.xml');
  const customFormats = new Map(children(first(styles.documentElement,'numFmts'),'numFmt').map(n=>[Number(n.getAttribute('numFmtId')),n.getAttribute('formatCode')]));
  const xfs = children(first(styles.documentElement,'cellXfs'),'xf');
  const isDateStyle = i => {
    const id = Number(xfs[Number(i)]?.getAttribute('numFmtId'));
    if ((id>=14&&id<=22)||(id>=45&&id<=47)) return true;
    const format=customFormats.get(id)||'';
    return /[dy]/i.test(format.replace(/"[^"]*"/g,''));
  };
  for (const node of children(first(book.documentElement, 'sheets'), 'sheet')) {
    const title = node.getAttribute('name');
    if (isInstructions(title)) continue;
    const target = map.get(node.getAttributeNS(R, 'id'));
    if (!target) throw Error(`La hoja ${title} no tiene relación en el libro.`);
    const path = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const doc = await loadXml(zip, path);
    sheets.push({node, doc, path, title, layout:sheetLayout(doc,strings)});
  }
  if (!sheets.length) throw Error('No se encontraron hojas con registros en el Excel.');
  const rows = [];
  for (const sheet of sheets) {
    const data = first(sheet.doc.documentElement, 'sheetData');
    for (const row of children(data, 'row')) {
      const number = Number(row.getAttribute('r'));
      if (number < sheet.layout.start || number >= sheet.layout.start+sheet.layout.capacity) continue;
      const cells = {};
      const values = {};
      for (const cell of children(row, 'c')) {
        const n = col((cell.getAttribute('r') || '').match(/^[A-Z]+/)?.[0] || '');
        if (n >= 1 && n <= 33) { cells[n] = cell; values[n] = cellValue(cell, strings); }
      }
      const plausibleName=Boolean(safeText(values[3])||safeText(values[4]));
      if (!isOrder(values[1]) && !(plausibleName && safeText(values[2]).length<=25)) continue;
      if (![2, 3, 4].some(n => safeText(values[n]))) continue;
      rows.push({id: `${sheet.title}!${number}`, sheet, row: number, cells, values});
    }
  }
  const dateStyles={};
  for(const c of [6,7,8,20]) {
    for(const sheet of sheets) {
      for(let n=sheet.layout.start;n<sheet.layout.start+sheet.layout.capacity;n++) {
        const cell=children(getRow(sheet.doc,n),'c').find(x=>x.getAttribute('r')===addr(c,n));
        if(cell && isDateStyle(cell.getAttribute('s'))) {dateStyles[c]=cell.getAttribute('s');break;}
      }
      if(dateStyles[c])break;
    }
  }
  dateStyles.fallback=String(xfs.findIndex((_,i)=>isDateStyle(i)));
  return {zip, book, rels, strings, sheets, rows, dateStyles, file, fileName:file.name || 'archivo.xlsx'};
}

function score(a, b) {
  const sameId = docKey(a.values[2]) && docKey(a.values[2]) === docKey(b.values[2]);
  const sameName = key(a.values[3]) && key(a.values[3]) === key(b.values[3]) && key(a.values[4]) === key(b.values[4]);
  const sameBirth = dateKey(a.values[6]) && dateKey(a.values[6]) === dateKey(b.values[6]);
  if (sameId) return 100;
  if (sameName && sameBirth) return 85;
  if (sameName) return 60;
  if (sameBirth && key(a.values[4]) && key(a.values[4]) === key(b.values[4])) return 45;
  return 0;
}

export function compare(previous, current) {
  const olds = previous.rows, news = current.rows;
  const count = arr => { const map = new Map(); for (const x of arr) { const d = docKey(x.values[2]); if (d) map.set(d, (map.get(d) || 0) + 1); } return map; };
  const oldCount = count(olds), newCount = count(news);
  const used = new Set(), automatic = new Map(), reviews = [];
  for (let i = 0; i < news.length; i++) {
    const b = news[i], doc = docKey(b.values[2]);
    const ranked = olds.map((a, j) => ({j, score:score(a,b)})).filter(x => x.score).sort((a,b) => b.score - a.score);
    const exact = ranked.filter(x => x.score === 100 && identity(olds[x.j]) === identity(b) && key(b.values[3]) && key(b.values[4]) && dateKey(b.values[6]) && isOrder(b.values[1]) && isOrder(olds[x.j].values[1]));
    if (doc && newCount.get(doc) === 1 && oldCount.get(doc) === 1 && exact.length === 1 && !used.has(exact[0].j)) {
      automatic.set(i, exact[0].j); used.add(exact[0].j);
    } else if (ranked.length || !doc || !key(b.values[3]) || !key(b.values[4]) || !dateKey(b.values[6]) || !isOrder(b.values[1])) {
      let reason = ranked.length ? (ranked[0].score === 100 ? 'Mismo documento, pero datos personales distintos o repetidos.' : 'Posible coincidencia sin documento idéntico.') : 'Faltan datos de identificación; confirma si es un ingreso.';
      if (doc && newCount.get(doc) > 1) reason = 'Documento repetido en el archivo actual; revisa ambas filas.';
      if (!isOrder(b.values[1])) reason = 'El número de orden no corresponde a una fila de niño; revisa la fila de origen.';
      reviews.push({index:i, ranked, reason});
    }
  }
  const assigned = new Map(automatic);
  // Conflicting exact matches never occur silently: those with duplicated NUIP are reviewed.
  return {previous, current, automatic, reviews, assigned};
}

export function consolidate(state) {
  const claimed = new Set(), rows = [];
  for (let i = 0; i < state.current.rows.length; i++) {
    const newRow = state.current.rows[i], oldIndex = state.assigned.get(i);
    if (oldIndex !== undefined && oldIndex !== null && oldIndex !== '') {
      if (claimed.has(Number(oldIndex))) throw Error('Una persona anterior fue asignada dos veces. Revisa las coincidencias.');
      claimed.add(Number(oldIndex));
      rows.push({status:'CONTINÚA', person:newRow, first:state.previous.rows[Number(oldIndex)], second:newRow, secondStart:hasMeasurement(newRow,20) ? 20 : 8});
    } else rows.push({status:'NUEVO',person:newRow,first:newRow,second:null});
  }
  for (let j = 0; j < state.previous.rows.length; j++) if (!claimed.has(j)) rows.push({status:'EGRESO',person:state.previous.rows[j],first:state.previous.rows[j],second:null});
  return sortConsolidated(rows);
}
function hasMeasurement(row, start) { return Array.from({length:12},(_,n) => safeText(row.values[start+n])).some(Boolean); }

export function outputValue(item,c) {
  if(Object.prototype.hasOwnProperty.call(item.overrides||{},c))return item.overrides[c];
  if(c<=7)return item.person.values[c]??'';
  if(c<=19)return item.first.values[c]??'';
  if(item.status==='EGRESO')return c===20?'EGRESO':'';
  if(!item.second)return '';
  return item.second.values[item.secondStart+c-20]??'';
}

export function sortConsolidated(rows) {
  const given=item=>key(outputValue(item,3));
  rows.sort((a,b)=>given(a).split(' ')[0].localeCompare(given(b).split(' ')[0],'es') || given(a).localeCompare(given(b),'es') || key(outputValue(a,4)).localeCompare(key(outputValue(b,4)),'es') || docKey(outputValue(a,2)).localeCompare(docKey(outputValue(b,2)),'es'));
  return rows;
}

export function pagePlan(current,count) {
  const sheets=current.sheets;
  const fullest=sheets.reduce((a,b)=>b.layout.capacity>a.layout.capacity?b:a);
  const plan=[];
  let remaining=count;
  do {
    const existing=sheets[plan.length];
    const template=existing && (existing.layout.capacity>=Math.min(20,remaining)||existing.layout.capacity>=fullest.layout.capacity)?existing:fullest;
    const take=Math.min(remaining,template.layout.capacity);
    plan.push({template,replaceExisting:Boolean(existing&&existing!==template),take});
    remaining-=take;
  }while(remaining>0);
  return plan;
}

function getRow(doc, number) {return children(first(doc.documentElement,'sheetData'),'row').find(n => Number(n.getAttribute('r')) === number);}
function getCell(row, index) {return children(row,'c').find(n => n.getAttribute('r') === addr(index,Number(row.getAttribute('r'))));}
function setCell(doc, row, index, source, literal) {
  let dest = getCell(row,index);
  if (!dest) {
    dest = doc.createElementNS(M,'c'); dest.setAttribute('r',addr(index,Number(row.getAttribute('r'))));
    const ref = getCell(getRow(doc,16),index);
    if (ref?.getAttribute('s')) dest.setAttribute('s',ref.getAttribute('s'));
    const next = children(row,'c').find(c => col((c.getAttribute('r')||'').match(/^[A-Z]+/)?.[0]||'') > index);
    row.insertBefore(dest,next || null);
  }
  while (dest.firstChild) dest.removeChild(dest.firstChild);
  dest.removeAttribute('t');
  if (literal !== undefined) {
    dest.setAttribute('t','inlineStr');
    const is = doc.createElementNS(M,'is'), t = doc.createElementNS(M,'t'); t.textContent=literal;is.appendChild(t);dest.appendChild(is);
  } else if (source) {
    const type = source.getAttribute('t'); if (type) dest.setAttribute('t',type);
    for (const ch of children(source)) if (ch.localName !== 'f') dest.appendChild(doc.importNode(ch,true));
    // Keep the destination's styles: style indices from the other workbook are unrelated.
  }
}

function overrideCell(doc,row,c,value,dateStyles){
  const raw=safeText(value);
  if(!raw){setCell(doc,row,c,null);return;}
  let numeric=null;
  if([6,7,8,20].includes(c)){
    const d=dateKey(raw),match=d.match(/^(\d{4})-(\d\d)-(\d\d)$/);
    if(match){
      const ms=Date.UTC(Number(match[1]),Number(match[2])-1,Number(match[3]));
      if(new Date(ms).toISOString().slice(0,10)===d)numeric=String(ms/86400000+25569);
    }
  }else if([9,10,11,13,17,21,22,23,25,29].includes(c)&&/^-?\d+(?:[.,]\d+)?$/.test(raw))numeric=raw.replace(',','.');
  if(numeric!==null){
    setCell(doc,row,c,null);
    const dest=getCell(row,c),v=doc.createElementNS(M,'v');v.textContent=numeric;dest.appendChild(v);
    if([6,7,8,20].includes(c)){
      const style=dateStyles[c]??dateStyles.fallback;
      if(style&&style!=='-1')dest.setAttribute('s',style);
    }
  }else setCell(doc,row,c,null,raw);
}

function putRow(doc, n, item, position, dateStyles) {
  const row = getRow(doc,n);
  if (!row) throw Error(`La plantilla no tiene la fila ${n}.`);
  for (let c=1;c<=31;c++) setCell(doc,row,c,null);
  const ordinal=doc.createElementNS(M,'v');ordinal.textContent=String(position);getCell(row,1).appendChild(ordinal);
  const copy=(c,source,sourceCol) => {
    const from=source.cells[sourceCol];
    if (from?.getAttribute('t')==='s') setCell(doc,row,c,null,String(source.values[sourceCol]??''));
    else {
      setCell(doc,row,c,from);
      if([6,7,8,20].includes(c) && from && !from.getAttribute('t') && /^\d+(?:\.\d+)?$/.test(safeText(source.values[sourceCol]))) {
        const dst=getCell(row,c),style=dateStyles[c] ?? dateStyles.fallback;
        if(style && style!=='-1')dst.setAttribute('s',style);
      }
    }
  };
  for (let c=2;c<=7;c++) copy(c,item.person,c);
  for (let c=8;c<=19;c++) copy(c,item.first,c);
  if (item.status === 'EGRESO') setCell(doc,row,20,null,'EGRESO');
  else if (item.second) for (let c=20;c<=31;c++) copy(c,item.second,item.secondStart+c-20);
  for(const [column,value] of Object.entries(item.overrides||{})){
    const c=Number(column);
    if(c>=2&&c<=31 && !(item.status==='EGRESO'&&c>=20))overrideCell(doc,row,c,value,dateStyles);
  }
}

async function detachSourceTables(sheet,zip,types){
  // The template's table definitions refer to its original header and row range.
  // Keeping them after rearranging rows makes Excel repair the workbook on open.
  const relationshipPath=sheet.path.replace('/worksheets/','/worksheets/_rels/')+'.rels';
  if(zip.file(relationshipPath)){
    const rels=await loadXml(zip,relationshipPath);
    for(const rel of children(rels.documentElement,'Relationship')){
      if(!rel.getAttribute('Type')?.endsWith('/table'))continue;
      const target=new URL(rel.getAttribute('Target'),'https://xlsx.local/'+sheet.path).pathname.slice(1);
      zip.remove(target);
      for(const override of children(types.documentElement,'Override')){
        if(override.getAttribute('PartName')==='/'+target)types.documentElement.removeChild(override);
      }
      rels.documentElement.removeChild(rel);
    }
    zip.file(relationshipPath,serializer.serializeToString(rels));
  }
  for(const part of children(sheet.doc.documentElement,'tableParts'))sheet.doc.documentElement.removeChild(part);
}

export async function makeOutput(state, consolidated) {
  // Start from a fresh copy so the user can download again without accumulating edits.
  const {zip,book,rels,sheets,dateStyles} = await readExcel(state.current.file);
  const plans = pagePlan({sheets},consolidated.length);
  const copies = plans.length;
  const listed = first(book.documentElement,'sheets');
  const relRoot = rels.documentElement;
  const types = await loadXml(zip,'[Content_Types].xml');
  for(const sheet of sheets)await detachSourceTables(sheet,zip,types);
  const existingIds = children(relRoot,'Relationship').map(r=>r.getAttribute('Id'));
  let nextRel = Math.max(0,...existingIds.map(s=>Number(s.match(/\d+$/)?.[0]||0)))+1;
  let nextSheet = Math.max(0,...children(listed,'sheet').map(s=>Number(s.getAttribute('sheetId')||0)))+1;
  let nextFile = Math.max(0,...Object.keys(zip.files).map(s=>Number(s.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1]||0)))+1;
  const existingNames = new Set(children(listed,'sheet').map(s=>s.getAttribute('name').toLocaleLowerCase('es-CO')));
  const originalCount = sheets.length;
  let lastNode=sheets[originalCount-1].node;
  const originals=new Map(sheets.map(s=>[s.path,serializer.serializeToString(s.doc)]));
  let offset=0;
  for (let i=0;i<copies;i++) {
    const plan=plans[i];
    let sheet;
    if (i<originalCount) {
      sheet=sheets[i];
      if(plan.replaceExisting){
        sheet.doc=xml(originals.get(plan.template.path));
        sheet.layout=plan.template.layout;
      }
    }
    else {
      const base = plan.template, fileNum=nextFile++, path=`xl/worksheets/sheet${fileNum}.xml`;
      const baseTitle=`Consolidado ${i+1}`;
      let title=baseTitle, suffix=2;
      while(existingNames.has(title.toLocaleLowerCase('es-CO'))) title=`${baseTitle} (${suffix++})`;
      existingNames.add(title.toLocaleLowerCase('es-CO'));
      const cloned=xml(originals.get(base.path));
      const node=book.createElementNS(M,'sheet'); node.setAttribute('name',title); node.setAttribute('sheetId',String(nextSheet++));node.setAttributeNS(R,'r:id',`rId${nextRel}`);
      listed.insertBefore(node,lastNode.nextSibling);lastNode=node;
      const rel=rels.createElementNS(P,'Relationship');rel.setAttribute('Id',`rId${nextRel++}`);rel.setAttribute('Type',R+'/worksheet');rel.setAttribute('Target',`worksheets/sheet${fileNum}.xml`);relRoot.appendChild(rel);
      const override=types.createElementNS(C,'Override');override.setAttribute('PartName','/'+path);override.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.appendChild(override);
      const baseRel=base.path.replace('/worksheets/','/worksheets/_rels/')+'.rels';
      if (zip.file(baseRel)) zip.file(`xl/worksheets/_rels/sheet${fileNum}.xml.rels`,await zip.file(baseRel).async('string'));
      sheet={doc:cloned,node,path,title,layout:base.layout};
    }
    const header=getRow(sheet.doc,sheet.layout.headerRow);
    if(header){
      for(const [c,label] of [[8,'TOMA Nº 1'],[20,'TOMA Nº 2']]){
        if(/TOMA/i.test(cellValue(getCell(header,c),state.current.strings)))setCell(sheet.doc,header,c,null,label);
      }
    }
    for (let j=0;j<sheet.layout.capacity;j++) {
      const item=j<plan.take?consolidated[offset+j]:null;
      const n=sheet.layout.start+j;
      if (item) putRow(sheet.doc,n,item,offset+j+1,dateStyles);
      else {const row=getRow(sheet.doc,n);if(row)for(let c=1;c<=31;c++)setCell(sheet.doc,row,c,null);}
    }
    offset+=plan.take;
    zip.file(sheet.path,serializer.serializeToString(sheet.doc));
  }
  for (let i=copies;i<originalCount;i++) {
    const sheet=sheets[i], relId=sheet.node.getAttributeNS(R,'id');
    listed.removeChild(sheet.node);
    const rel=children(relRoot,'Relationship').find(r=>r.getAttribute('Id')===relId);if(rel)relRoot.removeChild(rel);
    const override=children(types.documentElement,'Override').find(x=>x.getAttribute('PartName')==='/'+sheet.path);if(override)types.documentElement.removeChild(override);
    zip.remove(sheet.path);
  }
  zip.file('xl/workbook.xml',serializer.serializeToString(book));
  zip.file('xl/_rels/workbook.xml.rels',serializer.serializeToString(rels));
  zip.file('[Content_Types].xml',serializer.serializeToString(types));
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE'});
}

export {name, dateKey, safeText, score};
