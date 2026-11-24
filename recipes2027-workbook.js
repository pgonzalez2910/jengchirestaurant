/* Jeng Chi RecipeOS 2027 — workbook editing and printing layer.
   Source Excel snapshots are immutable. Only reviewed changes can be published. */
(function(){
'use strict';
var A={drafts:new Map(),published:new Map(),products:[],loading:null,selectedRow:null};
var UNITS=['oz','lb','fl oz','tsp','tbsp','cup','pt','qt','gal','each','piece','pcs','count','bottle','bag','box','case'];
var H=x=>String(x==null?'':x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
var bridge=()=>window.JCRecipeBridge;
var user=()=>bridge()?.getUser();
var client=()=>bridge()?.getClient();
var canEdit=()=>Boolean(user()&&['admin','editor'].includes(bridge()?.getRole()));
var fmt=v=>(v!==''&&v!==null&&v!==undefined&&Number.isFinite(Number(v))?Number(v).toFixed(3):String(v??''));
var k=(r,c)=>r+':'+c;
function original(sheet,row,col){
 var record=(sheet.rows||[]).find(x=>Number(x.r)===Number(row));
 var cell=record?.c?.find(x=>Number(x.c)===Number(col));
 return cell?.v??'';
}
function state(sheet){
 var p=A.published.get(sheet.component_id)||{cells:{},added_rows:[],removed_rows:[],version:0};
 return A.drafts.get(sheet.component_id)||p;
}
function cellValue(sheet,row,col){
 var item=state(sheet),key=k(row,col);
 return Object.prototype.hasOwnProperty.call(item.cells||{},key)?item.cells[key]:original(sheet,row,col);
}
function normalizeName(s){return String(s||'').toLowerCase().trim().replace(/\s+/g,' ');}
function lookupProduct(code){
 if(!code)return null;
 return A.products.find(p=>p.ct_code===code||(Array.isArray(p.previous_codes)&&p.previous_codes.includes(code)))||null;
}
function productForRow(sheet,row){
 var rawName=String(cellValue(sheet,row,2)||'');
 var code=String(cellValue(sheet,row,3)||'');
 var sel=String(cellValue(sheet,row,7)||'');
 if(/^(water|tap water|hot water|cold water)$/i.test(rawName.trim())) return {code:'R1326',name:'Tap Water',selector:'R1326 | Tap Water'};
 var parsed=sel.match(/^\s*([^|]+)\s*\|\s*(.+)$/);
 if(parsed){code=parsed[1].trim();rawName=parsed[2].trim();}
 var p=lookupProduct(code); if(p)return {code:p.ct_code,name:p.canonical_name,selector:p.ct_code+' | '+p.canonical_name};
 return {code:code,name:rawName,selector:sel||((code&&rawName)?code+' | '+rawName:'')};
}
function linkedFormulaValue(sheet,row,col,cell){
 if(col!==0||!cell?.f||typeof cell.v!=='number')return null;
 const refs=[...String(cell.f).matchAll(/'((?:[^']|'')+)'!\$?[AM]\$?(\d+)/g)];
 if(!refs.length)return null;
 const depName=refs[0][1].replace(/''/g,"'"),depRow=Number(refs[0][2]);
 if(refs.some(x=>x[1]!==refs[0][1]||Number(x[2])!==depRow))return null;
 const source=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.file===sheet.file&&x.sheet===depName);
 if(!source)return null;
 const input=(source.ingredients||[]).find(x=>Number(x.r)===depRow);if(!input)return null;
 const base=Number(input.q),now=Number(cellValue(source,depRow,0));
 if(!(base>0)||!Number.isFinite(now))return null;
 const current=state(sheet)?.cells||{};
 if(Object.prototype.hasOwnProperty.call(current,k(row,0)))return null;
 return Number((cell.v*now/base).toFixed(6));
}
function rendered(sheet,row,col,cell){
 var v=cellValue(sheet,row,col);
 var derived=linkedFormulaValue(sheet,row,col,cell);if(derived!==null)v=derived;
 var p=productForRow(sheet,row);
 if(col===7 && row>=7 && row<notesRow(sheet) && p.selector)v=p.selector;
 if(col===3 && p.code && (cell?.f||v===''||String(v)!==p.code))v=p.code;
 if(col===2 && p.name && (cell?.f||v===''||p.code==='R1326'))v=p.name;
 if(col===0 && (typeof v==='number'||v!==''&&/^[-+]?\d+(\.\d+)?$/.test(String(v))))v=fmt(v);
 if(col===6 && row===4 && typeof v==='number')v=fmt(v);
 return v??'';
}
async function load(){
 if(!user())return false;
 if(A.loading)return A.loading;
 A.loading=(async()=>{
  var uid=user().id, c=client();
  const [d,p,prod]=await Promise.all([
   c.from('jc_recipe_workbook_drafts').select('component_id,cells,added_rows,removed_rows,editor_id').eq('editor_id',uid),
   c.from('jc_recipe_workbook_publications').select('component_id,version,cells,added_rows,removed_rows'),
   c.from('jc_recipe_core_products').select('id,ct_code,canonical_name,code_status,previous_codes').order('canonical_name').limit(1500)
  ]);
  if(d.error)throw d.error;if(p.error)throw p.error;if(prod.error)throw prod.error;
  A.drafts=new Map((d.data||[]).map(x=>[x.component_id,x]));
  A.published=new Map((p.data||[]).map(x=>[x.component_id,x]));
  A.products=prod.data||[];
  return true;
 })().finally(()=>{A.loading=null;});
 return A.loading;
}
function sourceStatus(m){const e=document.getElementById('jeWorkbookSave');if(e)e.textContent=m;}
function toast(msg,error=false){
 if(window.JCRecipeCore?.toast)window.JCRecipeCore.toast(msg,error);
 else sourceStatus(msg);
}
function allStyles(sheet){return bridge()?.getLibrary()?.styles_by_file?.[sheet.file]||[];}
function isBlank(row){return !row?.c?.some(x=>String(x.v??'').trim()&&!/^\s*[-–]\s*$/.test(String(x.v??'')));}
function notesRow(sheet){
 return (sheet.rows||[]).find(row=>(row.c||[]).some(x=>/Toast Button Modifiers\s*\/\s*Notes/i.test(String(x.v||''))))?.r||sheet.max_row+1;
}
function allMerge(sheet){
 var starts=new Map(),covered=new Set();
 (sheet.merges||[]).forEach(m=>{
  var [c0,r0,c1,r1]=m;starts.set(k(r0,c0),[c1-c0+1,r1-r0+1]);
  for(var r=r0;r<=r1;r++)for(var c=c0;c<=c1;c++)if(r!==r0||c!==c0)covered.add(k(r,c));
 });
 return {starts,covered};
}
function rowHtml(sheet,rawRow,maxCol,merge,rowsExist,lastUsed,noteAt){
 var r=rawRow.r, cells=new Map((rawRow.c||[]).map(c=>[c.c,c]));
 var remove=(state(sheet).removed_rows||[]).some(x=>String(x)===String(r));
 var blank=isBlank(rawRow)&&!(state(sheet).cells&&Object.keys(state(sheet).cells).some(x=>x.startsWith(r+':')));
 var unused=(r>lastUsed+2&&r<noteAt&&blank);
 var cls=(remove?' jc-wb-removed':'')+(unused?' jc-wb-spare-row':'')+(blank&&r>=noteAt?' jc-wb-empty-note':'');
 var html='<tr class="jc-wb-row'+cls+'" data-row="'+r+'" style="height:'+Math.max(23,Math.min(100,Number(sheet.heights?.[r]||25)))*1.33+'px">';
 for(var col=0;col<=maxCol;col++){
  if(merge.covered.has(k(r,col)))continue;
  var cell=cells.get(col)||{c:col,v:''};var span=merge.starts.get(k(r,col));
  var picture=(col===5||col===6)&&r>=7&&r<noteAt;
  var editable=canEdit()&&!picture;
  var val=rendered(sheet,r,col,cell);
  var sty=allStyles(sheet)[cell.s||0]||''; if(col===7&&r>=6&&r<noteAt)sty+=';color:#17243a!important;background:#fff!important;text-align:left!important;white-space:normal!important;'; if(col===7&&r===6)val='PRODUCT SELECTOR · Code | Name';
  var isQty=(col===0 && (sheet.ingredients||[]).some(x=>x.r===r));
  var cls2='jc-wb-cell'+(editable?' jc-wb-editable':'')+(isQty?' jc-wb-qty':'');
  var title=(cell.f?'Original Excel formula: '+cell.f+' (changes are saved separately)':'');
  html+='<td class="'+cls2+'" data-r="'+r+'" data-c="'+col+'"'+(span?' colspan="'+span[0]+'" rowspan="'+span[1]+'"':'')+
    ' style="'+H(sty)+'" title="'+H(title)+'"><span class="jc-wb-value">'+H(val)+'</span>'+(editable&&(col===1||col===7)&&r<noteAt&&r>=7?'<span class="jc-wb-picker-caret" aria-hidden="true">▾</span>':'')+'</td>';
 }
 return html+'</tr>';
}
function addedHtml(sheet,item){
 const id=item.id,cl=item.cells||{},mark=allStyles(sheet),ing=(sheet.ingredients||[]).at(-1),ref=ing?(sheet.rows||[]).find(r=>r.r===ing.r):null;
 const map=new Map((ref?.c||[]).map(c=>[c.c,c]));
 var h='<tr class="jc-wb-row jc-wb-added" data-row="'+H(id)+'">';
 for(var c=0;c<8;c++){
  var style=mark[map.get(c)?.s||0]||'';
  var v=cl[c]??'';
  if(c===0&&v!==''&&Number.isFinite(Number(v)))v=fmt(v);
  h+='<td class="jc-wb-cell jc-wb-editable" data-r="'+H(id)+'" data-c="'+c+'" style="'+H(style)+'"><span class="jc-wb-value">'+H(v)+'</span></td>';
 }
 return h+'</tr>';
}
function render(sheet,targetId){
 var el=document.getElementById(targetId);if(!el||!sheet)return false;
 A.active=sheet;var heights=sheet.heights||{},widths=sheet.widths||{};
 var cols=Array.from({length:8},(_,i)=>Math.max(58,Math.min(440,Math.round((widths[i]||[14,10,31,19,16,16,18,42][i])*7))));
 var merge=allMerge(sheet),items=state(sheet),lastUsed=Math.max(7,...(sheet.ingredients||[]).map(x=>Number(x.r)||7)),noteAt=notesRow(sheet);
 var html='<table class="jc-excel-sheet jc-wb-advanced" aria-label="'+H(sheet.name)+' Excel worksheet"><colgroup>'+
 cols.map(w=>'<col style="width:'+w+'px">').join('')+'</colgroup><tbody>';
 var added=Array.isArray(items.added_rows)?items.added_rows:[];
 for(var row of sheet.rows||[]){
  html+=rowHtml(sheet,row,7,merge,true,lastUsed,noteAt);
  for(var x of added.filter(x=>Number(x.after)===Number(row.r)))html+=addedHtml(sheet,x);
 }
 html+='</tbody></table><datalist id="jcProductOptions">';
 for(var p of A.products)html+='<option value="'+H(p.ct_code+' | '+p.canonical_name)+'"></option>';
 html+='</datalist>';
 el.innerHTML=html;
 return true;
}
function currentDraft(sheet){
 var old=A.drafts.get(sheet.component_id)||A.published.get(sheet.component_id)||{};
 return {component_id:sheet.component_id,editor_id:user().id,cells:structuredClone(old.cells||{}),
  added_rows:structuredClone(old.added_rows||[]),removed_rows:[...(old.removed_rows||[])]};
}
async function save(sheet,obj){
 if(!canEdit())throw Error('Authorized editing session required');
 const payload={...obj,updated_at:new Date().toISOString()};
 const {error}=await client().from('jc_recipe_workbook_drafts').upsert(payload,{onConflict:'component_id,editor_id'});
 if(error)throw error;
 A.drafts.set(sheet.component_id,payload);
 sourceStatus('Saved to Supabase · unpublished draft');
 return true;
}
function qtyFor(sheet,row){
 var q=cellValue(sheet,row,0);return Number.isFinite(Number(q))?Number(q):null;
}
function unitFor(sheet,row){return String(cellValue(sheet,row,1)||'').trim();}
function groupConvert(q,from,to,name){
 from=String(from||'').toLowerCase().trim();to=String(to||'').toLowerCase().trim();
 if(from===to)return {quantity:q,estimated:false};
 const mass={oz:1,lb:16},vol={'fl oz':1,tsp:1/6,tbsp:.5,cup:8,pt:16,qt:32,gal:128};
 const pieces={each:1,piece:1,pcs:1,count:1};
 for(var table of [mass,vol,pieces]){
  if(from in table&&to in table)return {quantity:q*table[from]/table[to],estimated:false};
 }
 // Water only: approximate room-temperature density, not a generic mass/volume conversion.
 if(/^(tap water|water|hot water|cold water)$/i.test(name||'')){
  const densityLbPerGal=8.345;
  if((from in mass)&&(to in vol))return {quantity:(q*mass[from]/16/densityLbPerGal*128)/vol[to],estimated:true};
  if((from in vol)&&(to in mass))return {quantity:(q*vol[from]/128*densityLbPerGal*16)/mass[to],estimated:true};
 }
 return null;
}
async function setCell(sheet,row,col,value){
 var obj=currentDraft(sheet),key=k(row,col);
 const originalValue=cellValue(sheet,row,col);
 if(String(value)===String(originalValue))return true;
 if(col===0&&value!==''&&(!Number.isFinite(Number(value))||Number(value)<0))throw Error('Ingredient quantity must be a nonnegative number');
 if(typeof row==='string'&&row.startsWith('new-')){
  const target=obj.added_rows.find(x=>x.id===row);
  if(!target)throw Error('Added ingredient row no longer exists');
  target.cells=target.cells||{};target.cells[col]=(col===0&&value!==''?Number(value):value);if(col===7&&value.includes(' | ')){const [code,...name]=value.split(' | ');target.cells[3]=code;target.cells[2]=name.join(' | ');}
 }else{
  obj.cells[key]=(col===0&&value!==''?Number(value):value);
  if(col===7&&String(value).includes(' | ')){
   const parts=String(value).split(' | '),code=parts.shift().trim(),productName=parts.join(' | ').trim();
   const actual=lookupProduct(code);
   if(actual){obj.cells[k(row,2)]=actual.canonical_name;obj.cells[k(row,3)]=actual.ct_code;obj.cells[key]=actual.ct_code+' | '+actual.canonical_name;}
   else if(productName){obj.cells[k(row,2)]=productName;obj.cells[k(row,3)]=code;}
  }
  if(col===1){
   const originalUnit=unitFor(sheet,row),q=qtyFor(sheet,row),p=productForRow(sheet,row);
   if(q!==null){
    const cv=groupConvert(q,originalUnit,value,p.name);
    if(cv){obj.cells[k(row,0)]=Number(cv.quantity.toFixed(6));if(cv.estimated)sourceStatus('Water conversion is approximate; verify measured density and actual yield');}
   }
  }
 }
 await save(sheet,obj);
 return true;
}
async function addRow(sheet){
 var max=Math.max(7,...(sheet.ingredients||[]).map(x=>Number(x.r)||7));
 var obj=currentDraft(sheet);
 obj.added_rows.push({id:'new-'+Date.now()+'-'+Math.random().toString(36).slice(2,9),after:max,cells:{0:'',1:'oz',2:'',3:'',4:'',5:'',6:'',7:''}});
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');
 return true;
}
async function removeRow(sheet,rowKey){
 if(rowKey==null)throw Error('First select the row you want to remove');
 if(!confirm('Remove this worksheet row from the working draft? Your original Excel workbook is protected.'))return false;
 var obj=currentDraft(sheet);
 if(String(rowKey).startsWith('new-'))obj.added_rows=obj.added_rows.filter(x=>x.id!==rowKey);
 else if(!obj.removed_rows.map(String).includes(String(rowKey)))obj.removed_rows.push(Number(rowKey));
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');A.selectedRow=null;return true;
}
async function restoreRows(sheet){
 var obj=currentDraft(sheet);obj.removed_rows=[];
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');
}
async function publish(sheet){
 if(bridge()?.getRole()!=='admin')throw Error('Owner approval is required');
 if(!A.drafts.has(sheet.component_id))throw Error('No unpublished changes to publish');
 const count=(bridge()?.getLibrary()?.sheets||[]).filter(x=>x.component_id===sheet.component_id).length;
 if(!confirm('Publish this master component to '+count+' matching worksheet'+(count!==1?'s':'')+'? This will approve shared changes.'))return;
 const {data,error}=await client().rpc('jc_recipe_publish_workbook_component',{p_component_id:sheet.component_id});
 if(error)throw error;
 await load();sourceStatus('Published master component · version '+data);
 render(sheet,'jeWorkbookGrid');
}
async function scaleBatch(sheet){
 if(!canEdit())throw Error('Authorized editor required');
 const from=Number(String(cellValue(sheet,4,6)||sheet.yield||'').match(/\d+(?:\.\d+)?/)?.[0]||0);
 if(!(from>0))throw Error('A numeric source target yield is required before scaling');
 const unit=String(cellValue(sheet,4,6)||sheet.yield||'').replace(/^\s*[\d,.]+\s*/,'').trim();
 const result=prompt('New target batch yield'+(unit?' ('+unit+')':'')+':',fmt(from));
 if(result===null)return;
 const to=Number(result);
 if(!(to>0)||!Number.isFinite(to))throw Error('Enter a positive target yield');
 const factor=to/from;
 if(!confirm('Scale '+sheet.name+' by '+factor.toFixed(3)+' ×? Ingredient quantities change proportionally. Actual finished yield must still be measured.'))return;
 const obj=currentDraft(sheet);
 for(const ingredient of sheet.ingredients||[]){
  const row=ingredient.r;
  const old=Number(cellValue(sheet,row,0));
  if(Number.isFinite(old))obj.cells[k(row,0)]=Number((old*factor).toFixed(6));
 }
 for(const extra of obj.added_rows||[]){
  const q=Number(extra.cells?.[0]);
  if(Number.isFinite(q)&&extra.cells?.[0]!=='')extra.cells[0]=Number((q*factor).toFixed(6));
 }
 obj.cells[k(4,6)]=unit?to.toFixed(3)+' '+unit:to;
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');
 sourceStatus('Target yield and ingredient quantities scaled together · actual measured yield remains separate');
}
async function changeCode(){
 if(bridge()?.getRole()!=='admin')throw Error('Owner approval is required');
 const old=prompt('Current CrunchTime product code:');if(!old)return;
 const next=prompt('New approved product code for '+old+':');if(!next)return;
 if(!confirm('Approve replacing '+old+' with '+next+' across recipe product selectors? The previous code will remain in the audit history.'))return;
 const {error}=await client().rpc('jc_recipe_update_product_code',{p_old:old.trim(),p_new:next.trim()});
 if(error)throw error;await load();if(A.active)render(A.active,'jeWorkbookGrid');
 sourceStatus('Central product code updated · '+old+' → '+next);
}
function print(){
 document.body.classList.add('je-print-workbook');
 window.addEventListener('afterprint',()=>document.body.classList.remove('je-print-workbook'),{once:true});
 window.print();
}
function editCell(td){
 if(!canEdit()||!td.classList.contains('jc-wb-editable')||td.querySelector('input,select,textarea'))return;
 const sheet=A.active;if(!sheet)return;
 const row=td.dataset.r.startsWith('new-')?td.dataset.r:Number(td.dataset.r),col=Number(td.dataset.c);
 const old=String(col===7&&!String(row).startsWith('new-')?productForRow(sheet,row).selector:row.toString().startsWith('new-')?(state(sheet).added_rows||[]).find(x=>x.id===row)?.cells?.[col]||'':cellValue(sheet,row,col));
 var control;
 if(col===7){
  control=document.createElement('select');
  const values=[...new Set(['',old,...A.products.map(p=>p.ct_code+' | '+p.canonical_name)])];
  values.forEach(v=>{const opt=document.createElement('option');opt.value=v;opt.textContent=v||'Select a product…';control.appendChild(opt)});control.value=old;
 }else if(col===1){
  control=document.createElement('select');
  UNITS.forEach(u=>{const opt=document.createElement('option');opt.value=u;opt.textContent=u;control.appendChild(opt)});
  if(old&&!UNITS.includes(old)){const opt=document.createElement('option');opt.value=old;opt.textContent=old;control.prepend(opt)}
  control.value=old;
 }else if((col===1||col===0||col===7)||old.length<70){
  control=document.createElement('input');control.type='text';control.value=col===0&&old!==''?fmt(old):old;
  if(col===7)control.setAttribute('list','jcProductOptions');
 }else{
  control=document.createElement('textarea');control.rows=2;control.value=old;
 }
 control.className='jc-wb-edit-control';control.setAttribute('aria-label','Edit worksheet cell');
 td.replaceChildren(control);control.focus();if(control.select)control.select();if(control.tagName==='SELECT'&&control.showPicker){try{control.showPicker()}catch(_){}}
 let committed=false;
 async function complete(commit=true){
  if(committed)return;committed=true;
  try{if(commit)await setCell(sheet,row,col,control.value)}catch(e){sourceStatus('Not saved: '+e.message)}
  render(sheet,'jeWorkbookGrid');
 }
 control.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();complete()}if(e.key==='Escape'){e.preventDefault();complete(false)}});
 control.addEventListener('blur',()=>complete());
 if(control.tagName==='SELECT')control.addEventListener('change',()=>complete());
}
document.addEventListener('dblclick',e=>{
 const td=e.target.closest?.('#jeWorkbookGrid td.jc-wb-editable');
 if(td)editCell(td);
});
document.addEventListener('click',e=>{
 const td=e.target.closest?.('#jeWorkbookGrid tr[data-row]');
 if(td){document.querySelectorAll('#jeWorkbookGrid tr.jc-wb-selected').forEach(x=>x.classList.remove('jc-wb-selected'));td.classList.add('jc-wb-selected');A.selectedRow=td.dataset.row;}
 const picker=e.target.closest?.('#jeWorkbookGrid td.jc-wb-editable');
 if(picker&&[1,7].includes(Number(picker.dataset.c))&&!picker.querySelector('input,select,textarea'))editCell(picker);
});
window.JCRecipeAdvanced={
 load,render,addRow,removeRow,restoreRows,publish,changeCode,scaleBatch,print,groupConvert,
 get selectedRow(){return A.selectedRow},
 get active(){return A.active},
 get productCount(){return A.products.length},
 get version(){return A.published.get(A.active?.component_id)?.version||0}
};
})();