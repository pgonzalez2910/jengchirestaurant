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
 sheet=canonicalSheet(sheet);
 var record=(sheet.rows||[]).find(x=>Number(x.r)===Number(row));
 var cell=record?.c?.find(x=>Number(x.c)===Number(col));
 return cell?.v??'';
}
const approvedMushroomMasters={"DM-PREP-INTERNAL":{"name":"Dried Mushroom Prep","master":"2aae8bf9-10fb-5051-8e1a-a2fece6d31cc","ids":["2aae8bf9-10fb-5051-8e1a-a2fece6d31cc","b0978670-2650-516c-9111-7e58a290e024","14a4b515-ceda-5c3f-be31-def83e73fefb","00343328-cfd5-58bb-b52f-c2de52ed09f7","44d2072f-2321-5799-bebf-620c88d5889a","af6d1eaa-251d-5f7f-b74d-dc45af81e8a5"],"defaults":{"4:6":"15 lb","8:0":2,"8:1":"lb","8:2":"Produce Mushroom Dried","8:3":"P6758","8:7":"P6758 | Produce Mushroom Dried","9:0":10,"9:1":"qt","9:2":"Tap Water","9:3":"R1326","9:7":"R1326 | Tap Water"}},"FM-PREP-INTERNAL":{"name":"Fungus Mushroom Prep","master":"63854625-ea09-5853-be55-e46f077cd0f3","ids":["63854625-ea09-5853-be55-e46f077cd0f3","a381962a-6796-59f4-8ed7-4bddd27e03fe","69b40037-1d21-5de2-b201-5c9618a3ced1","d71c68d9-dcc6-5f33-9033-69dcfb1d629e"],"defaults":{"4:6":"15 lb","8:0":2.5,"8:1":"lb","8:2":"Mushroom Fungus Strips","8:3":"P2874","8:7":"P2874 | Mushroom Fungus Strips","9:0":18,"9:1":"qt","9:2":"Tap Water","9:3":"R1326","9:7":"R1326 | Tap Water"}},"SL-INTERNAL":{"name":"Special Lard","master":"4593d49a-4de7-5b2c-94ba-d75738596152","ids":["4593d49a-4de7-5b2c-94ba-d75738596152","c4453b8d-c554-5515-9137-54443c69fe00","36c6cd6b-b059-5aa5-b467-1d7938c23914","6d8cd551-2f20-5005-9069-18a6ccd5f727"],"defaults":{"4:6":"90 lb (5 batches)"}},"PK-J-INTERNAL":{"name":"Pork Jelly","master":"b7d4bffe-ad3a-5584-a923-27739b5efefa","ids":["b7d4bffe-ad3a-5584-a923-27739b5efefa","ab76cd2d-8ae9-54eb-8b4e-7c89a8c8194b"],"defaults":{"4:6":"35 lb","16:0":3.5,"16:1":"bottle","16:2":"Gelatin Mix Unflavored","16:3":"P6712","16:7":"P6712 | Gelatin Mix Unflavored"}},"DD-INTERNAL":{"name":"White Dumpling Dough","master":"e08a0d11-ee85-5a69-85b0-6ade9e8cecd7","ids":["e08a0d11-ee85-5a69-85b0-6ade9e8cecd7","45e82113-9e38-54df-8625-ac3551bc7018","24b1125f-e083-59c0-9375-5584533dce7c","6090d151-088c-5798-b64e-59a0df220f4e","0077b05e-ae47-5e49-b2f3-f2594eaaebd5","10a88043-3446-553d-9ca3-7d0a339dda6a","9d57db67-cfa9-51f0-ae5f-9d901b762c5d","f67bed53-ae93-5313-bbc3-243fb4e2a904","2eef8213-695e-5054-ac3f-f4f1fbcc3863","2fb05b4c-c6ae-5db5-87a7-4a975e97377d","07ba04fd-0213-5165-95b2-0a9d1cf214ee"],"defaults":{"4:6":"95 lb","8:0":54,"8:1":"lb","8:2":"Flour All Purpose","8:3":"P1167","8:7":"P1167 | Flour All Purpose","9:0":18,"9:1":"lb","9:2":"Special Lard","9:3":"SL-INTERNAL","9:7":"SL-INTERNAL | Special Lard","10:0":13.76,"10:1":"qt","10:2":"Tap Water","10:3":"R1326","10:7":"R1326 | Tap Water"}},"CKOIL-INTERNAL":{"name":"Chicken Oil","master":"6b1038d1-4556-5f9e-a1bd-42c37603e464","ids":["6b1038d1-4556-5f9e-a1bd-42c37603e464","906070b0-e186-5dd6-bf6e-bea9d395ce8d","019c3df2-e651-5546-9f53-a455db9e1283"],"defaults":{"4:6":"6 lb","8:0":20,"8:1":"lb","8:2":"Chicken Breast","8:3":"P1101","8:7":"P1101 | Chicken Breast","9:0":2,"9:1":"lb","9:2":"Produce Onion Yellow","9:3":"P1448","9:7":"P1448 | Produce Onion Yellow"}}};
function mushroomMaster(sheet){return Object.values(approvedMushroomMasters).find(g=>g.ids.includes(sheet?.component_id));}
function canonicalSheet(sheet){const g=mushroomMaster(sheet);return g?(bridge()?.getLibrary()?.sheets||[]).find(x=>x.component_id===g.master)||sheet:sheet;}
function state(sheet){
 const g=mushroomMaster(sheet);
 if(g){const p=A.published.get(g.master)||{},d=A.drafts.get(g.master)||p;return {...d,cells:{...g.defaults,...(d.cells||{}),'2:2':g.name},added_rows:d.added_rows||[],removed_rows:d.removed_rows||[]};}

 var p=A.published.get(sheet.component_id)||{cells:{},added_rows:[],removed_rows:[],version:0};
 return A.drafts.get(sheet.component_id)||p;
}
// Active recipe books have empty Notes/Prep sections. Immutable source snapshots
// retain original text and registered identifiers for audit and dependency resolution.
const blankNoteCache=new WeakMap();
const noteHeaderText=v=>/^\s*(?:Toast Button Modifiers\s*\/\s*Notes|Prep\s*\/\s*Plating Instructions)\s*:?\s*$/i.test(String(v??''));
const noteMarker=v=>/^\s*[*]+\s*$/.test(String(v??''));
const separatePanel=v=>/^(?:\s*(?:Linked (?:Portions|Recipes|Sub-Recipes)|CrunchTime Registration|Chocolate Rolls\s*[—–-]\s*Production Yield|Production Yield\s*\/|Portion Yield|Stock Yield|Actual Yield|Target Yield|Yield\s*\/\s*Portion|Orders per (?:Actual|Target)|Portion Size|Portions\s*\/\s*Batch|Cost Control|Used In\s*\/\s*Return To|Return To\s*$|Recipe Reference|Open\s+|Return to\s+|Large Noodle Portions|Pork Chop Breading))/i.test(String(v??''));
function blankNoteCells(sheet){
 const src=canonicalSheet(sheet);
 if(blankNoteCache.has(src))return blankNoteCache.get(src);
 const keys=new Set();let section=false;
 const merges=src.merges||[];
 for(const row of [...(src.rows||[])].sort((a,b)=>Number(a.r)-Number(b.r))){
  const rr=Number(row.r),cells=row.c||[],values=cells.filter(c=>String(c.v??'').trim()!=='');
  if(cells.some(c=>noteHeaderText(c.v))){
   section=true;
   for(const m of merges.filter(m=>Number(m[1])===rr&&Number(m[3])===rr&&Number(m[0])<=1&&Number(m[2])>=5)){
    for(let c=Number(m[0]);c<=Number(m[2]);c++)if(!cells.some(x=>Number(x.c)===c&&noteHeaderText(x.v)))keys.add(k(rr,c));
   }
   continue;
  }
  if(!section)continue;
  const marker=cells.some(c=>Number(c.c)===0&&noteMarker(c.v));
  const wide=merges.filter(m=>Number(m[1])===rr&&Number(m[3])===rr&&Number(m[0])<=1&&Number(m[2])>=5);
  const substantive=values.filter(c=>!noteMarker(c.v));
  if(!marker&&substantive.some(c=>separatePanel(c.v))){section=false;continue;}
  // A separate numeric grid is not a note section.
  if(!marker&&!wide.length&&substantive.some(c=>Number(c.c)>=2||typeof c.v==='number'||/^\s*\d+(?:\.\d+)?\s*$/.test(String(c.v)))){section=false;continue;}
  if(marker){keys.add(k(rr,1));for(const m of wide)if(Number(m[0])!==0)keys.add(k(rr,Number(m[0])));}
  else if(wide.length){for(const m of wide)for(let c=Number(m[0]);c<=Number(m[2]);c++)keys.add(k(rr,c));}
  else if(substantive.length<=1){keys.add(k(rr,0));keys.add(k(rr,1));}
  // Preserve the asterisk itself even if an unusual source merge covers column A.
  if(marker)keys.delete(k(rr,0));
 }
 blankNoteCache.set(src,keys);return keys;
}
function isBlankNoteCell(sheet,row,col){return blankNoteCells(sheet).has(k(row,col));}
function cellValue(sheet,row,col){
 if(isBlankNoteCell(sheet,row,col))return '';
 if(sheet.variantServing){if(isYield(sheet,row,col))return sheet.variantServing;if(Number(row)===2&&Number(col)===2)return sheet.name;if(Number(row)===4&&Number(col)===0)return '';}
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
 if(code==='DD-INTERNAL'&&(bridge()?.getLibrary()?.sheets||[]).some(x=>x.file===sheet.file&&x.component_id==='62423561-d4a2-5384-98b9-7acc25bbca72'))return {code,name:rawName,selector:code+' | '+rawName};
 if(approvedMushroomMasters[code]){rawName=approvedMushroomMasters[code].name;return {code,name:rawName,selector:code+' | '+rawName};}
 var p=lookupProduct(code); if(p&&normalizeName(p.canonical_name)===normalizeName(rawName))return {code:p.ct_code,name:rawName,selector:p.ct_code+' | '+rawName};
 return {code:code,name:rawName,selector:(code&&rawName)?code+' | '+rawName:''};
}
let protectedMain=null;
function isProtected(sheet){return !!sheet&&protectedMain===sheet.id;}
const yieldCache=new WeakMap();
function yieldAddress(sheet){
 if(yieldCache.has(sheet))return yieldCache.get(sheet);
 let address=null;
 for(const row of sheet.rows||[]){
  const label=(row.c||[]).find(c=>/^\s*yield\s*:?\s*$/i.test(String(c.v??'')));
  if(!label)continue;
  const nr=Number(row.r)+1,nc=Number(label.c);
  const merged=(sheet.merges||[]).find(m=>nr>=m[1]&&nr<=m[3]&&nc>=m[0]&&nc<=m[2]);
  address=merged?{row:Number(merged[1]),col:Number(merged[0])}:{row:nr,col:nc};break;
 }
 if(address&&!/^\s*\d/.test(String(original(sheet,address.row,address.col)))) {
 const row=(sheet.rows||[]).find(x=>Number(x.r)===address.row);
 const candidates=(row?.c||[]).filter(x=>Number(x.c)>=address.col&&Number(x.c)<=7&&/^\s*\d/.test(String(x.v??'')));
 if(candidates.length===1)address={row:address.row,col:Number(candidates[0].c)};
 }
 if(!address)address={row:4,col:6};
 yieldCache.set(sheet,address);return address;
}
function isYield(sheet,row,col){const a=yieldAddress(sheet);return Number(row)===a.row&&Number(col)===a.col;}
function yieldValue(sheet,originalOnly=false){const a=yieldAddress(sheet);const raw=originalOnly?original(sheet,a.row,a.col):cellValue(sheet,a.row,a.col);const n=Number(String(raw||sheet.yield||'').match(/^\s*(\d+(?:\.\d+)?)/)?.[1]);return n>0?n:null;}
// Direct per-serving links use live batch quantities and the explicit visible piece yield.
// Archived helper columns M/N are not present in library snapshots.
const explicitPorkJellyLinks={
 'b1400339-c678-5491-905e-1c2d42be51c1':{row:9,src:'dceab539-4783-59eb-9998-86666be3acb2',sourceRow:9},
 'beaf9807-9bdc-55f2-a59f-59f58c7b61f4':{row:10,src:'f8424828-bb4d-521b-afa3-253e45c7dac8',sourceRow:10},
 'b502e50d-4cdd-5caa-9dd4-443596fc1cf6':{row:14,src:'4347a594-ce3e-5532-99e9-b53a44d46af7',sourceRow:13},
 'c9341297-6c4e-571e-854f-390ae3dbe12e':{row:11,src:'e16eb7a1-94dd-5452-9cd9-5ad9008c6f44',sourceRow:11}
};
function fillingReference(sheet,row){
 const pinned=explicitPorkJellyLinks[sheet.component_id];
 if(pinned&&Number(row)===pinned.row){
  const src=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.component_id===pinned.src&&x.file===sheet.file);
  if(!src)throw Error('Linked filling unavailable for '+sheet.name);
  return {src,row:pinned.sourceRow};
 }
 const cell=(sheet.rows||[]).find(x=>Number(x.r)===Number(row))?.c?.find(x=>Number(x.c)===0);
 const f=String(cell?.f||'');
 const direct=f.match(/'((?:[^']|'')+)'!\$?M\$?(\d+)\s*\*\s*\$?N\$?3\s*\/\s*'((?:[^']|'')+)'!\$?N\$?3/i);
 const indexed=f.match(/INDEX\(\s*'((?:[^']|'')+)'!\$?M\$?(\d+):\$?M\$?(\d+)\s*,\s*MATCH\(\s*(\d+)\s*,\s*'((?:[^']|'')+)'!\$?O\$?\d+:\$?O\$?\d+\s*,\s*0\s*\)\s*\)\s*\*\s*\$?N\$?3\s*\/\s*'((?:[^']|'')+)'!\$?N\$?3/i);
 if(!direct&&!indexed)return null;
 if(direct&&direct[1]!==direct[3]||indexed&&(indexed[1]!==indexed[5]||indexed[1]!==indexed[6]))return null;
 const depName=(direct||indexed)[1].replace(/''/g,"'");
 const src=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.file===sheet.file&&x.sheet===depName);
 if(!src||!/\bfilling\b/i.test(String(src.name)+' '+String(src.sheet))||!dumplingCountYield(src,cellValue(src,yieldAddress(src).row,yieldAddress(src).col)))return null;
 let rr=direct?Number(direct[2]):null;
 if(indexed){
  const candidates=(src.ingredients||[]).filter(i=>Number(i.r)>=Number(indexed[2])&&Number(i.r)<=Number(indexed[3])).sort((a,b)=>Number(a.r)-Number(b.r));
  const item=candidates.find(i=>Number(i.i)===Number(indexed[4])-1)||candidates[Number(indexed[4])-1];
  const target=(sheet.ingredients||[]).find(i=>Number(i.r)===Number(row));
  if(!item||target?.code&&item.code!==target.code)return null;
  rr=Number(item.r);
 }
 return {src,row:rr};
}
// Approved yield-linked consumption policy: preserve each parent's share of the original 35 lb Pork Jelly batch.
function preparedQuantityFactor(code){
 if(String(code).trim().toUpperCase()!=='PK-J-INTERNAL')return 1;
 const g=approvedMushroomMasters['PK-J-INTERNAL'];
 const src=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.component_id===g.master);
 const stored=A.drafts.get(g.master)||A.published.get(g.master);
 const raw=stored?.cells?.['4:6']??g.defaults['4:6'];
 const y=Number(String(raw).match(/^\s*(\d+(?:\.\d+)?)/)?.[1]);
 if(!(y>0))throw Error('Invalid saved Pork Jelly master yield');
 return y/35;
}
function linkedPreparedQuantity(q,code){
 return q!==''&&q!=null&&Number.isFinite(Number(q))?Number(q)*preparedQuantityFactor(code):q;
}
function fillingLine(sheet,src,row,servings,outputUnit){
 const count=yieldValue(src);if(!(count>0&&servings>0))return null;
 const st=state(src);
 let q,u,p;
 if(typeof row==='string'&&row.startsWith('new-')){
  const v=(st.added_rows||[]).find(x=>x.id===row)?.cells||{};
  q=v[0];u=String(v[1]||'');p={code:String(v[3]||''),name:String(v[2]||''),selector:String(v[7]||'')};
 }else{
  if((st.removed_rows||[]).map(String).includes(String(row)))return {q:'',unit:'',code:'',name:'',selector:''};
  q=cellValue(src,row,0);u=String(cellValue(src,row,1)||'');p=productForRow(src,row);
 }
 if(q===''||q==null||!p.code)return {q:'',unit:'',code:'',name:'',selector:''};
 if(!Number.isFinite(Number(q)))return null;
 let amount=linkedPreparedQuantity(q,p.code)*servings/count,unit=u;
 if(outputUnit&&u){const cv=groupConvert(amount,u,outputUnit,p.name);if(cv){amount=cv.quantity;unit=outputUnit;}}
 return {q:amount,unit,code:p.code,name:p.name,selector:p.code+' | '+p.name};
}
function liveFillingLine(sheet,row,servings=yieldValue(sheet)){
 const ref=fillingReference(sheet,row);if(!ref)return null;
 return fillingLine(sheet,ref.src,ref.row,servings,String(original(sheet,row,1)||'oz'));
}
function linkedFormulaValue(sheet,row,col,cell){
 if(col!==0||!cell?.f||typeof cell.v!=='number')return null;
 // A finished dough mass is not a number of wrappers. Preserve archived per-order dough weights until a measured wrapper standard is available.
 if(String(original(sheet,row,3)).trim()==='DD-INTERNAL'&&/SUM\s*\(/i.test(cell.f))return cell.v;
 // Finished dough is now measured in pounds; a pounds value is not a piece denominator.
 // Archived per-serving wrapper consumption is retained until an actual wrapper weight is specified.
 const doughRef=String(cell.f).match(/'((?:[^']|'')+)'!\$?N\$?3/i);
 if(doughRef&&/SUM\(/i.test(cell.f)){
  const dep=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.file===sheet.file&&x.sheet===doughRef[1].replace(/''/g,"'"));
  if(dep&&mushroomMaster(dep)===approvedMushroomMasters['DD-INTERNAL'])return cell.v;
 }
 const refs=[...String(cell.f).matchAll(/'((?:[^']|'')+)'!\$?[AM]\$?(\d+)/g)];
 if(!refs.length)return null;
 const depName=refs[0][1].replace(/''/g,"'"),depRow=Number(refs[0][2]);
 if(refs.some(x=>x[1]!==refs[0][1]||Number(x[2])!==depRow))return null;
 const source=(bridge()?.getLibrary()?.sheets||[]).find(x=>x.file===sheet.file&&x.sheet===depName);
 if(!source)return null;
 const input=(source.ingredients||[]).find(x=>Number(x.r)===depRow);if(!input)return null;
 const base=Number(input.q),raw=linkedPreparedQuantity(cellValue(source,depRow,0),productForRow(source,depRow).code);
 if(raw===''||raw==null)return 0;
 const now=Number(raw);
 if(!(base>0)||!Number.isFinite(now))return null;
 let converted=now;
 const oldUnit=String(input.u||''),newUnit=String(cellValue(source,depRow,1)||oldUnit);
 if(normalizeName(oldUnit)!==normalizeName(newUnit)){
  const cv=groupConvert(now,newUnit,oldUnit,productForRow(source,depRow).name);
  if(!cv)return null;converted=cv.quantity;
 }
 let batchFactor=1;
 if(/!\$?N\$?3\b/i.test(String(cell.f))&&String(cell.f).includes('/')){
  const oldYield=yieldValue(source,true),newYield=yieldValue(source);
  if(oldYield&&newYield)batchFactor=oldYield/newYield;
 }
 const current=state(sheet)?.cells||{};
 if(!isProtected(sheet)&&Object.prototype.hasOwnProperty.call(current,k(row,0)))return null;
 return Number((cell.v*converted/base*batchFactor).toFixed(6));
}
function dumplingCountYield(sheet,raw){
 if(!/^\s*\d+(?:\.\d+)?\s*(?:pcs?|pieces?|dumplings?)?\s*$/i.test(String(raw??'')))return false;
 if(sheet.variantServing)return true;
 if(!/\bdumplings?\b/i.test(String(sheet.file||'')))return false;
 const name=String(sheet.name||'')+' '+String(sheet.sheet||'');
 if(/\b(dough|jelly|stock|broth|oil|lard|grease|prep|trim|blend|sauce)\b/i.test(name))return false;
 return isProtected(sheet)||/\bfilling\b/i.test(name)||/\bdumplings?\b/i.test(name);
}
function rendered(sheet,row,col,cell){
 if(isBlankNoteCell(sheet,row,col))return '';
 const live=ingredientRow(sheet,row)?liveFillingLine(sheet,row):null;
 if(live&&[0,1,2,3,7].includes(col))return col===0?(live.q===''?'':fmt(live.q)):col===1?live.unit:col===2?live.name:col===3?live.code:live.selector;
 var v=cellValue(sheet,row,col);
 var derived=linkedFormulaValue(sheet,row,col,cell);if(derived!==null)v=derived;
 if(derived===null&&col===0&&ingredientRow(sheet,row))v=linkedPreparedQuantity(v,productForRow(sheet,row).code);
 if(isProtected(sheet)&&ingredientRow(sheet,row)&&col===0){const base=yieldValue(sheet,true),target=yieldValue(sheet);const q=derived??linkedPreparedQuantity(original(sheet,row,col),productForRow(sheet,row).code);if(base&&target&&q!==''&&q!=null&&Number.isFinite(Number(q)))v=Number(q)*target/base;}
 if(ingredientRow(sheet,row)&&[2,3,7].includes(col)){
 const p=productForRow(sheet,row);
 if(col===7)v=p.selector||'';
 if(col===3&&p.code)v=p.code;
 if(col===2&&p.name)v=p.name;
 }
 if(col===0 && ingredientRow(sheet,row) && (typeof v==='number'||v!==''&&/^[-+]?\d+(\.\d+)?$/.test(String(v))))v=fmt(v);
 if(isYield(sheet,row,col)&&dumplingCountYield(sheet,v))v=Number(String(v).match(/^[\s]*\d+(?:\.\d+)?/)[0]).toLocaleString('en-US')+' dumplings';
 else if(isYield(sheet,row,col)&&typeof v==='number')v=fmt(v);
 return v??'';
}
async function load(){
 if(!user())return false;
 if(A.loading)return A.loading;
 A.loading=(async()=>{
  var uid=user().id, c=client();
  const [d,p,prod]=await Promise.all([
   c.from('jc_recipe_workbook_drafts').select('component_id,cells,added_rows,removed_rows,editor_id,updated_at').eq('editor_id',uid),
   c.from('jc_recipe_workbook_publications').select('component_id,version,cells,added_rows,removed_rows'),
   c.from('jc_recipe_core_products').select('id,ct_code,canonical_name,code_status,previous_codes').order('canonical_name').limit(1500)
  ]);
  if(d.error)throw d.error;if(p.error)throw p.error;if(prod.error)throw prod.error;
  const incoming=new Map((d.data||[]).map(x=>[x.component_id,x]));
  for(const [id,old] of A.drafts)if(old.updated_at&&(!incoming.get(id)?.updated_at||old.updated_at>incoming.get(id).updated_at))incoming.set(id,old);
  A.drafts=incoming;
  A.published=new Map((p.data||[]).map(x=>[x.component_id,x]));
  A.products=prod.data||[];
  return true;
 })().finally(()=>{A.loading=null;});
 return A.loading;
}
function sourceStatus(m){const e=document.getElementById('jeWorkbookSave');if(!e)return;const kind=/not saved|⚠|error|failed/i.test(m)&&!/^(✓|Published)/.test(m)?'failed':/saving|⏳/i.test(m)?'saving':/^(✓|Published)|saved|updated/i.test(m)?'saved':'ready';e.className='jc-autosave '+kind;const icon=document.createElement('i');icon.className='fa-solid '+(kind==='saving'?'fa-cloud-arrow-up fa-fade':kind==='saved'?'fa-circle-check':kind==='failed'?'fa-circle-exclamation':'fa-cloud');icon.setAttribute('aria-hidden','true');e.replaceChildren(icon,document.createTextNode(' '+String(m).replace(/^[⏳✓⚠]+\s*/u,'')));if(kind==='saved')e.title='Saved at '+new Date().toLocaleTimeString();else e.title='';}
function toast(msg,error=false){
 if(window.JCRecipeCore?.toast)window.JCRecipeCore.toast(msg,error);
 else sourceStatus(msg);
}
function allStyles(sheet){return bridge()?.getLibrary()?.styles_by_file?.[sheet.file]||[];}
function isBlank(row){return !row?.c?.some(x=>String(x.v??'').trim()&&!/^\s*[-–]\s*$/.test(String(x.v??'')));}
function notesRow(sheet){
 return (sheet.rows||[]).find(row=>(row.c||[]).some(x=>/Toast Button Modifiers\s*\/\s*Notes/i.test(String(x.v||''))))?.r||sheet.max_row+1;
}
const boundsCache=new WeakMap();
function ingredientBounds(sheet){
 if(boundsCache.has(sheet))return boundsCache.get(sheet);
 const header=(sheet.rows||[]).find(r=>(r.c||[]).some(c=>Number(c.c)===0&&/^\s*(qty|quantity)\s*$/i.test(String(c.v)))&&(r.c||[]).some(c=>Number(c.c)===1&&/units?/i.test(String(c.v))))?.r;
 const first=Math.min(...(sheet.ingredients||[]).map(x=>Number(x.r)).filter(Number.isFinite));
 const start=header!=null?Number(header)+1:Number.isFinite(first)?first:Infinity;
 const end=(sheet.rows||[]).find(r=>Number(r.r)>=start&&(r.c||[]).some(c=>/Toast Button Modifiers|Prep\s*\/\s*Plating Instructions/i.test(String(c.v||''))))?.r??notesRow(sheet);
 const result={start,end:Number(end)};boundsCache.set(sheet,result);return result;
}
function ingredientRow(sheet,row){if(String(row).startsWith('new-'))return true;const b=ingredientBounds(sheet);return Number(row)>=b.start&&Number(row)<b.end;}
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
 var remove=!isProtected(sheet)&&(state(sheet).removed_rows||[]).some(x=>String(x)===String(r));
 var blank=isBlank(rawRow)&&!(state(sheet).cells&&Object.keys(state(sheet).cells).some(x=>x.startsWith(r+':')));
 var unused=(r>lastUsed+2&&r<noteAt&&blank);
 var cls=(remove?' jc-wb-removed':'')+(unused?' jc-wb-spare-row':'')+(blank&&r>=noteAt?' jc-wb-empty-note':'');
 var html='<tr class="jc-wb-row'+cls+'" data-row="'+r+'" style="height:'+Math.max(23,Math.min(100,Number(sheet.heights?.[r]||25)))*1.33+'px">';
 for(var col=0;col<=maxCol;col++){
  if(merge.covered.has(k(r,col)))continue;
  var cell=cells.get(col)||{c:col,v:''};var span=merge.starts.get(k(r,col));
  var isIngredient=ingredientRow(sheet,r),header=Number(r)===ingredientBounds(sheet).start-1;
  var picture=(col===5||col===6)&&isIngredient;
  var editable=canEdit()&&!picture&&!header&&(isIngredient?[0,1,4,7].includes(col):col!==7);
 if(isProtected(sheet))editable=canEdit()&&(isYield(sheet,r,col)||(!isIngredient&&Number(r)>noteAt&&col===1&&!/Prep\s*\/\s*Plating Instructions/i.test(String(cell.v||''))));
  if(sheet.variantServing||isBlankNoteCell(sheet,r,col))editable=false;
  var val=rendered(sheet,r,col,cell);
  var sty=allStyles(sheet)[cell.s||0]||''; if(col===7&&(isIngredient||header))sty+=';color:#17243a!important;background:'+(header?'#a3ce68':'#fff')+'!important;text-align:left!important;white-space:normal!important;'; if(col===7&&header)val='PRODUCT SELECTOR · Code | Name';
  if(isIngredient&&[0,1,2,3,7].includes(col)&&fillingReference(sheet,r))editable=false;
 var isQty=(col===0 && (sheet.ingredients||[]).some(x=>x.r===r));
  var cls2='jc-wb-cell'+(editable?' jc-wb-editable':'')+(isQty?' jc-wb-qty':'')+(isYield(sheet,r,col)&&editable?' jc-wb-yield-edit':'');
  var title=(!isBlankNoteCell(sheet,r,col)&&cell.f?'Original Excel formula: '+cell.f+' (changes are saved separately)':'');
  html+='<td class="'+cls2+'" data-r="'+r+'" data-c="'+col+'"'+(span?' colspan="'+span[0]+'" rowspan="'+span[1]+'"':'')+
    ' style="'+H(sty)+'" title="'+H(title)+'">'+(editable&&isYield(sheet,r,col)&&(isProtected(sheet)||isPorkJelly(sheet))?'<input class="jc-direct-yield'+(isPorkJelly(sheet)?' jc-batch-yield':'')+'" type="'+(isPorkJelly(sheet)?'text':'number')+'" inputmode="decimal" min="'+(isPorkJelly(sheet)?'0.001':'1')+'" step="'+(isPorkJelly(sheet)?'any':'1')+'" aria-label="'+(isPorkJelly(sheet)?'Target batch yield in pounds':'Serving yield in pieces')+'" value="'+H(yieldValue(sheet)||'')+'">'+(isPorkJelly(sheet)?'<span class="jc-yield-unit">lb</span>':dumplingCountYield(sheet,cellValue(sheet,r,col))?'<span class="jc-yield-unit">dumplings</span>':''):'<span class="jc-wb-value">'+H(val)+'</span>')+(editable&&isIngredient&&(col===1||col===7)?'<span class="jc-wb-picker-caret" aria-hidden="true">▾</span>':'')+'</td>';
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
  if(c===0&&v!==''&&Number.isFinite(Number(v)))v=fmt(linkedPreparedQuantity(v,cl[3]));
  h+='<td class="jc-wb-cell'+(canEdit()&&[0,1,4,7].includes(c)?' jc-wb-editable':'')+'" data-r="'+H(id)+'" data-c="'+c+'" style="'+H(style)+'"><span class="jc-wb-value">'+H(v)+'</span></td>';
 }
 return h+'</tr>';
}
function extraFillingLines(sheet){
 const sources=new Map();
 for(const r of sheet.rows||[]){const ref=fillingReference(sheet,r.r);if(ref){if(!sources.has(ref.src.id))sources.set(ref.src.id,{src:ref.src,mapped:new Set()});sources.get(ref.src.id).mapped.add(ref.row);}}
 const extra=[];
 for(const {src,mapped} of sources.values()){
  const candidates=[...(src.rows||[]).filter(r=>ingredientRow(src,r.r)).map(r=>Number(r.r)),...(state(src).added_rows||[]).map(r=>r.id)];
  for(const rr of candidates){if(mapped.has(rr))continue;const x=fillingLine(sheet,src,rr,yieldValue(sheet),'oz');if(x&&x.q!==''&&x.code)extra.push({...x,key:src.id+':'+rr});}
 }
 return extra;
}
function extraFillingHtml(sheet,rows){
 if(!rows.length)return '';
 return rows.map(x=>'<tr class="jc-wb-row jc-live-extra"><td data-live-extra="'+H(x.key)+'">'+H(fmt(x.q))+'</td><td>'+H(x.unit)+'</td><td>'+H(x.name)+'</td><td>'+H(x.code)+'</td><td></td><td colspan="2"></td><td>'+H(x.selector)+'</td></tr>').join('');
}
function liveIngredientLines(sheet){
 const out=[],st=state(sheet),removed=new Set((st.removed_rows||[]).map(String));
 for(const row of sheet.rows||[]){
  const rr=Number(row.r);if(!ingredientRow(sheet,rr)||removed.has(String(rr)))continue;
  const live=liveFillingLine(sheet,rr);
  if(live){if(live.q!==''&&live.code)out.push(live);continue;}
  const p=productForRow(sheet,rr),cell=(row.c||[]).find(c=>Number(c.c)===0);
  let q=linkedFormulaValue(sheet,rr,0,cell)??linkedPreparedQuantity(cellValue(sheet,rr,0),p.code);
  if(isProtected(sheet)){const base=yieldValue(sheet,true),target=yieldValue(sheet);if(base&&target&&q!==''&&q!=null&&Number.isFinite(Number(q)))q=Number(q)*target/base;}
  if(q!==''&&q!=null&&Number.isFinite(Number(q))&&p.code)out.push({...p,q:Number(q),unit:String(cellValue(sheet,rr,1)||'')});
 }
 for(const row of st.added_rows||[]){const v=row.cells||{};if(v[0]!==''&&v[0]!=null&&Number.isFinite(Number(v[0]))&&v[3])out.push({q:linkedPreparedQuantity(v[0],v[3]),unit:String(v[1]||''),code:String(v[3]),name:String(v[2]||'')});}
 out.push(...extraFillingLines(sheet));
 return out;
}
function preparedSource(sheet,code){
 const reg=registry(),key=String(code||'').trim().toUpperCase(),all=(reg.byCode.get(key)||[]).filter(x=>x.component_id!==sheet.component_id);
 const local=all.filter(x=>x.file===sheet.file);
 const candidates=local.length?local:all;
 if(!candidates.length)return null;
 const roots=new Map(candidates.map(x=>[mushroomMaster(x)?.master||x.component_id,x]));
 if(roots.size===1)return canonicalSheet([...roots.values()][0]);
 const first=candidates[0],sig=signature(first);
 if(sig&&candidates.every(x=>signature(x)===sig))return canonicalSheet(first);
 return null;
}
function masterYieldLabel(sheet,row){
 const live=liveFillingLine(sheet,row),p=live||productForRow(sheet,row);
 if(!p.code||!/-INTERNAL$/i.test(p.code))return '';
 const source=preparedSource(sheet,p.code);if(!source)return '';
 const y=measuredYield(source);if(!y)return '';
 const master=mushroomMaster(source)?.master||source.component_id;
 return 'Master: '+fmt(y.q)+' '+y.unit+(p.code==='PK-J-INTERNAL'?' · yield-linked':'')+(A.drafts.has(master)?' · draft':' · published/source');
}
function measuredYield(sheet){
 const a=yieldAddress(sheet),raw=String(cellValue(sheet,a.row,a.col)||'').trim();
 const m=raw.match(/^(\d+(?:\.\d+)?)\s*(lb|lbs|oz|fl oz|qt|qtt|qts|gal|cups?|pt|each|pcs?|pieces?|dumplings?|portions?|count)\b/i);
 if(!m||!(Number(m[1])>0)){
  // Archived dough counts identify wrappers, not finished weight.
  if(/^\s*\d+(?:\.\d+)?\s*$/.test(raw)&&/\bdough\b/i.test(sheet.name+' '+sheet.sheet))return {q:Number(raw),unit:'each'};
  return null;
 }
 const u=m[2].toLowerCase(),unit=/^(?:pcs?|pieces?|dumplings?|portions?|count)$/.test(u)?'each':u==='lbs'?'lb':/^(qtt|qts)$/.test(u)?'qt':u==='cups'?'cup':u;
 return {q:Number(m[1]),unit};
}
function livePrepBreakdown(sheet){
 const groups=[],warnings=[],links=new Set(),resolved=new Set(),issues=new Map();
 function issue(line,reason,src=null){
  const key=line.code+'|'+(src?.component_id||'')+'|'+reason;
  issues.set(key,{code:line.code,name:line.name,reason,source:src?.sheet||'',portion:line.q,unit:line.unit});
 }
 function visit(parent,line,factor,path,depth){
  if(depth>8||! /-INTERNAL$/i.test(line.code)||!(line.q>0))return;
  const src=preparedSource(parent,line.code);
  if(!src){issue(line,'Master recipe not uniquely identified');return;}
  const id=mushroomMaster(src)?.master||src.component_id;
  links.add(id);
  if(path.includes(id)){issue(line,'Circular dependency',src);return;}
  const y=measuredYield(src),cv=y?groupConvert(line.q*factor,line.unit,y.unit,line.name):null;
  if(!y){issue(line,'Finished batch yield is missing or has no usable unit',src);return;}
  if(!cv){issue(line,'Yield is in '+y.unit+' but portion is in '+line.unit+'; compatible finished weight or volume required',src);return;}
  const ratio=cv.quantity/y.q,lines=liveIngredientLines(src).filter(x=>x.q!==0);
  resolved.add(id);
  groups.push({code:line.code,name:mushroomMaster(src)?.name||src.name,portion:line.q*factor,unit:line.unit,yield:y,depth,lines:lines.map(x=>({...x,q:x.q*ratio}))});
  for(const child of lines)visit(src,child,ratio,[...path,id],depth+1);
 }
 for(const line of liveIngredientLines(sheet))visit(sheet,line,1,[sheet.component_id],0);
 for(const x of issues.values())warnings.push(x.name+' — '+x.reason);
 return {groups,warnings,issues:[...issues.values()],linkedCount:links.size,resolvedCount:resolved.size};
}
function breakdownAmount(q,unit){
 // Display small allocations in practical units, leaving stored recipe data untouched.
 const u=String(unit||'').toLowerCase().trim();
 const to=u==='lb'&&Math.abs(q)<1?'oz':(u==='qt'||u==='gal'||u==='cup'||u==='pt')&&Math.abs(q)<1?'fl oz':null;
 if(to){const cv=groupConvert(q,u,to,'');if(cv)return {q:cv.quantity,unit:to};}
 return {q,unit};
}
function showLiveBreakdown(sheet){
 const grid=document.getElementById('jeWorkbookGrid');if(!grid)return;
 let panel=document.getElementById('jcLiveBreakdown');
 const result=livePrepBreakdown(sheet);
 if(!result.groups.length&&!result.issues.length){panel?.remove();return;}
 if(!panel){panel=document.createElement('details');panel.id='jcLiveBreakdown';panel.className='jc-dependency-report jc-live-breakdown';grid.closest('.je-wb-scroll')?.before(panel);}
 panel.innerHTML='<summary>Live prepared-ingredient breakdown · '+result.resolvedCount+' calculated'+(result.issues.length?' · '+result.issues.length+' need yield or link review':'')+'</summary><div class="jc-dependency-body"><p>Raw ingredient allocation for the selected recipe. Not additional inventory consumption.</p>'+
 result.groups.map(g=>'<section class="jc-live-prep"><div class="jc-prep-heading"><strong>'+H(g.name)+'</strong><span>'+H(fmt(g.portion)+' '+g.unit)+' used · Batch yield '+H(fmt(g.yield.q)+' '+g.yield.unit)+'</span></div><table><thead><tr><th>Raw / supporting ingredient</th><th>Quantity</th><th>Unit</th></tr></thead><tbody>'+g.lines.map(x=>{const v=breakdownAmount(x.q,x.unit);return '<tr><td>'+H(x.name)+'</td><td>'+H(fmt(v.q))+'</td><td>'+H(v.unit)+'</td></tr>';}).join('')+'</tbody></table></section>').join('')+
 (result.issues.length?'<div class="jc-prep-review"><strong>Items requiring review</strong>'+result.issues.map(x=>'<div><b>'+H(x.name)+'</b> · '+H(x.code)+'<span>'+H(x.reason)+(x.source?' · Linked worksheet: '+H(x.source):'')+'</span></div>').join('')+'</div>':'')+'</div>';
}
function render(sheet,targetId){
 const master=canonicalSheet(sheet);if(master!==sheet){sheet={...master,id:sheet.id,component_id:sheet.component_id,file:master.file,sheet:sheet.sheet,name:mushroomMaster(sheet).name};}
 var el=document.getElementById(targetId);if(!el||!sheet)return false;
 flushBatchInput=null;A.active=sheet;A.selectedCell=null;var heights=sheet.heights||{},widths=sheet.widths||{};
 var cols=Array.from({length:8},(_,i)=>Math.max(58,Math.min(440,Math.round((widths[i]||[14,10,31,19,16,16,18,42][i])*7))));
 var merge=allMerge(sheet),items=state(sheet),lastUsed=Math.max(7,...(sheet.ingredients||[]).map(x=>Number(x.r)||7)),noteAt=notesRow(sheet);
 var html='<table class="jc-excel-sheet jc-wb-advanced" aria-label="'+H(sheet.name)+' Excel worksheet"><colgroup>'+
 cols.map(w=>'<col style="width:'+w+'px">').join('')+'</colgroup><tbody>';
 var added=Array.isArray(items.added_rows)?items.added_rows:[];
 const liveExtra=extraFillingLines(sheet);let insertedLive=false;
 for(var row of sheet.rows||[]){
  if(!insertedLive&&Number(row.r)>=noteAt){html+=extraFillingHtml(sheet,liveExtra);insertedLive=true;}
  html+=rowHtml(sheet,row,7,merge,true,lastUsed,noteAt);
  for(var x of added.filter(x=>Number(x.after)===Number(row.r)))if(!isProtected(sheet))html+=addedHtml(sheet,x);
 }
 if(!insertedLive)html+=extraFillingHtml(sheet,liveExtra);
 html+='</tbody></table>';
 el.innerHTML=html;
 if(isPorkJelly(sheet))bindBatchYieldInput(sheet,el.querySelector('.jc-batch-yield'));else bindServingInput(sheet,el.querySelector('.jc-direct-yield'));
 showDependencyReport(sheet,A.lastSync?.id===sheet.id?A.lastSync.ids:null);
 showLiveBreakdown(sheet);
 return true;
}
function bindBatchYieldInput(sheet,input){
 if(!input)return;
 let timer=null,flight=null;
 const read=()=>{const m=String(input.value).trim().match(/^(\d+(?:\.\d+)?)\s*(?:lbs?)?$/i);return m&&Number(m[1])>0?Number(m[1]):null;};
 function preview(draft){
  if(A.active!==sheet)return;
  for(const node of document.querySelectorAll('#jeWorkbookGrid td[data-c="0"] .jc-wb-value')){
   const rr=node.closest('td').dataset.r;
   const q=rr.startsWith('new-')?draft.added_rows?.find(x=>x.id===rr)?.cells?.[0]:draft.cells[k(Number(rr),0)];
   if(q!==undefined&&q!==null&&q!=='')node.textContent=fmt(q);
  }
 }
 function restore(){const n=yieldValue(sheet);input.value=n||'';if(n)preview(scaledBatchDraft(sheet,n));}
 async function commit(){
  clearTimeout(timer);
  if(flight){await flight;return commit();}
  const n=read();if(n===null){restore();throw Error('Enter a positive target yield in pounds');}
  if(n===yieldValue(sheet)){preview(scaledBatchDraft(sheet,n));return;}
  const obj=scaledBatchDraft(sheet,n);preview(obj);
  flight=save(sheet,obj);
  try{await flight;}catch(e){restore();throw e;}finally{flight=null;}
  if(A.active===sheet){
   const next=read();if(next&&next!==n)return commit();
   input.value=n;preview(obj);showLiveBreakdown(sheet);
  }
 }
 const auto=()=>commit().catch(e=>sourceStatus('Not saved: '+e.message));
 flushBatchInput=commit;
 input.addEventListener('focus',()=>input.select());
 input.addEventListener('input',()=>{clearTimeout(timer);const n=read();if(n===null)return;preview(scaledBatchDraft(sheet,n));sourceStatus('⏳ Recalculating batch · saving automatically');timer=setTimeout(auto,500);});
 input.addEventListener('change',auto);input.addEventListener('blur',auto);
 input.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();auto();}if(e.key==='Escape'&&!flight){clearTimeout(timer);restore();}});
}
function bindServingInput(sheet,input){
 if(!input)return;
 let committed=yieldValue(sheet),pending=false;
 function preview(n){
  const base=yieldValue(sheet,true);if(!(base>0&&Number.isInteger(n)&&n>0))return;
  for(const node of document.querySelectorAll('#jeWorkbookGrid td[data-c="0"] .jc-wb-value')){
   const rr=Number(node.closest('td').dataset.r);if(!ingredientRow(sheet,rr))continue;
   const sourceCell=(sheet.rows||[]).find(x=>Number(x.r)===rr)?.c?.find(x=>Number(x.c)===0);
   const live=liveFillingLine(sheet,rr,n);if(live){node.textContent=live.q===''?'':fmt(live.q);continue;}
   const q=linkedFormulaValue(sheet,rr,0,sourceCell)??original(sheet,rr,0);
   if(q!==''&&q!=null&&Number.isFinite(Number(q)))node.textContent=fmt(Number(q)*n/base);
  }
 }
 async function commit(){
  if(pending)return;
  const n=Number(input.value);
  if(!Number.isInteger(n)||n<1){input.value=committed||'';preview(committed);sourceStatus('Enter a positive whole-number serving yield');return;}
  if(n===committed)return;
  pending=true;input.disabled=true;
  try{const a=yieldAddress(sheet);await setCell(sheet,a.row,a.col,n);committed=n;preview(n);showLiveBreakdown(sheet);}
  catch(err){input.value=committed||'';preview(committed);sourceStatus('Not saved: '+err.message);}
  finally{pending=false;input.disabled=false;}
 }
 input.addEventListener('focus',()=>input.select());
 input.addEventListener('input',()=>preview(Number(input.value)));
 input.addEventListener('change',commit);
 input.addEventListener('blur',commit);
 input.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();commit();}if(e.key==='Escape'){input.value=committed||'';preview(committed);}});
}
// Internal-code registry: explicit registration takes precedence; never infer a code from
// an ambiguous global product name. Local ingredient references can identify a unique master.
let registryCache=null,registryLibrary=null;
function registry(){
 const lib=bridge()?.getLibrary();if(registryLibrary===lib&&registryCache)return registryCache;
 const sheets=(lib?.sheets||[]).filter(s=>s.component_id&&!/^(database\d*|lookup lists|product match|lookups)$/i.test(s.sheet||''));
 const codeOf=new Map(),byCode=new Map(),byFile=new Map();
 const normCode=x=>String(x||'').trim().toUpperCase();
 const internal=x=>/[-_]INTERNAL$/i.test(String(x||''));
 for(const s of sheets){if(!byFile.has(s.file))byFile.set(s.file,[]);byFile.get(s.file).push(s);}
 for(const s of sheets){
  const codes=new Set();
  for(const r of s.rows||[])for(const v of r.c||[]){
   const match=String(v.v??'').match(/\b(?:CT\s+)?internal\s+(?:product\s+)?code\s*:\s*([A-Z0-9_-]+)\b/i);
   if(match&&internal(match[1]))codes.add(normCode(match[1]));
  }
  if(!codes.size){
   for(const parent of byFile.get(s.file)||[])if(parent.id!==s.id)for(const i of parent.ingredients||[]){
    if(internal(i.code)&&normalizeName(i.name)===normalizeName(s.name||s.sheet)&&
       (byFile.get(s.file)||[]).filter(x=>normalizeName(x.name||x.sheet)===normalizeName(i.name)).length===1)codes.add(normCode(i.code));
   }
  }
  const approved=Object.entries(approvedMushroomMasters).find(([key,g])=>g.ids.includes(s.component_id));if(approved){codes.clear();codes.add(approved[0]);}
  if(codes.size===1){const code=[...codes][0];codeOf.set(s.id,code);if(!byCode.has(code))byCode.set(code,[]);byCode.get(code).push(s);}
 }
 registryLibrary=lib;return registryCache={sheets,codeOf,byCode,byFile};
}
function originalLines(sheet){
 return (sheet.ingredients||[]).filter(i=>i.code&&i.q!==''&&Number.isFinite(Number(i.q))&&ingredientRow(sheet,i.r))
 .map(i=>({row:Number(i.r),code:String(i.code).toUpperCase(),q:Number(i.q),unit:normalizeName(i.u)}));
}
function signature(sheet){
 const lines=originalLines(sheet);
 // Require exact ingredient quantities, codes, row addresses and yield. Missing/ambiguous units are not merged.
 const y=String(sheet.yield??'').trim().toLowerCase().replace(/\s+/g,' ');
 if(!lines.length||!y||!yieldValue(sheet,true))return null;
 return JSON.stringify([normalizeName(sheet.name||sheet.sheet),y,lines,yieldAddress(sheet)]);
}
function matchingMasters(sheet){
 const reg=registry(),code=reg.codeOf.get(sheet.id),group=reg.byCode.get(code)||[];
 const approved=mushroomMaster(sheet);if(approved)return {code,conflicts:group.filter(x=>!approved.ids.includes(x.component_id)),matches:group.filter(x=>approved.ids.includes(x.component_id))};
 const sig=signature(sheet);
 const conflicts=group.filter(s=>s.component_id!==sheet.component_id&&(!sig||signature(s)!==sig));
 const matches=group.filter(s=>s.component_id===sheet.component_id||(sig&&signature(s)===sig));
 return {code,conflicts,matches};
}
function dependencyReport(sheet){
 const reg=registry(),group=matchingMasters(sheet),visited=new Set(),found=new Map(),ambiguous=new Map();
 const queue=[...new Map([sheet,...group.matches].map(x=>[x.id,{sheet:x,depth:0}])).values()];
 while(queue.length){
  const {sheet:child,depth}=queue.shift();if(visited.has(child.id))continue;visited.add(child.id);
  const code=reg.codeOf.get(child.id);
  const peers=matchingMasters(child);
  const valid=new Set([child.id,...peers.matches.map(x=>x.id)]);
  for(const parent of reg.sheets){
   if(parent.id===child.id||parent.component_id===child.component_id)continue;
   const raw=state(parent);
   const ing=(parent.ingredients||[]).filter(i=>ingredientRow(parent,i.r)&&!(raw.removed_rows||[]).map(String).includes(String(i.r)));
   const codes=[...ing.map(i=>liveFillingLine(parent,i.r)?.code||productForRow(parent,i.r).code),...(raw.added_rows||[]).map(x=>x.cells?.[3])];
   const byCode=!!code&&codes.some(x=>String(x||'').trim().toUpperCase()===code);
   const refs=parent.file===child.file&&(parent.rows||[]).some(r=>(r.c||[]).some(c=>String(c.f||'').toLowerCase().includes(("'"+String(child.sheet).replace(/'/g,"''")+"'!").toLowerCase())));
   if(!byCode&&!refs)continue;
   if(!refs&&byCode){
    const local=(reg.byFile.get(parent.file)||[]).filter(x=>x.id!==parent.id&&reg.codeOf.get(x.id)===code);
    if(local.length){if(local.every(x=>!valid.has(x.id)))continue;if(local.some(x=>!valid.has(x.id))){ambiguous.set(parent.id,parent);continue;}}
    else if(peers.conflicts.length){ambiguous.set(parent.id,parent);continue;}
   }
   if(!found.has(parent.id))found.set(parent.id,{sheet:parent,depth:depth+1,kind:refs?'Formula dependency — evaluate on open':'Consumes master output — portion retained'});
   if(!visited.has(parent.id))queue.push({sheet:parent,depth:depth+1});
  }
 }
 const roots=new Set([sheet,...group.matches].map(x=>x.id));
 const confirmed=[...found.values()].filter(x=>!roots.has(x.sheet.id)&&x.sheet.component_id!==sheet.component_id);
 return {group,found:confirmed,ambiguous:[...ambiguous.values()].filter(x=>!found.has(x.id)&&!roots.has(x.id))};
}
function showDependencyReport(sheet,savedIds=null){
 if(sheet.variantServing)return;
 if(A.active!==sheet)return;
 const grid=document.getElementById('jeWorkbookGrid');if(!grid)return;
 let panel=document.getElementById('jcDependencyReport');
 if(!panel){panel=document.createElement('details');panel.id='jcDependencyReport';panel.className='jc-dependency-report';grid.closest('.je-wb-scroll')?.before(panel);}
 const {group,found,ambiguous}=dependencyReport(sheet);
 const ids=new Set(savedIds||[]),copies=group.matches.filter(x=>x.id!==sheet.id);const approved=mushroomMaster(sheet);
 panel.innerHTML='<summary>Recipe dependencies · '+H(group.code||sheet.name)+' · '+found.length+' verified linked worksheet'+(found.length===1?'':'s')+(group.conflicts.length?' · ⚠ Code conflict':'')+'</summary>'+
 '<div class="jc-dependency-body">'+(savedIds?'<p>Saved '+ids.size+' component draft'+(ids.size===1?'':'s')+'. Original workbooks unchanged.</p>':'')+
 copies.map(x=>'<div>✓ '+H(x.file)+' / '+H(x.sheet)+' — '+(ids.has(approved?.master||x.component_id)?'shared master draft saved':approved?'shared canonical master':'matching master')+'</div>').join('')+
 group.conflicts.map(x=>'<div>⚠ '+H(x.file)+' / '+H(x.sheet)+' — different source formula or yield; not merged</div>').join('')+
 found.map(x=>'<div>'+H(x.sheet.file)+' / '+H(x.sheet.sheet)+' — '+H((x.depth===1?'Direct: ':'Indirect: ')+x.kind)+'</div>').join('')+ (ambiguous.length?'<p>Unresolved code references ('+ambiguous.length+') — excluded from verified dependencies:</p>'+ambiguous.map(x=>'<div>⚠ '+H(x.file)+' / '+H(x.sheet)+' — source not uniquely identified</div>').join(''):'')+
 (group.code==='DD-INTERNAL'?'<p>Shared finished dough yield is measured in pounds. Original per-dish wrapper portions are retained; pounds are never used as a piece count. Purple dough is a separate formulation.</p>':'')+(group.code==='PK-J-INTERNAL'?'<p>Pork Jelly portions are linked to master yield.</p>':'')+'</div>';
}
function currentDraft(sheet){
 const g=mushroomMaster(sheet);if(g)sheet=canonicalSheet(sheet);
 var old=g?state(sheet):A.drafts.get(sheet.component_id)||A.published.get(sheet.component_id)||{};
 return {component_id:sheet.component_id,editor_id:user().id,cells:structuredClone(old.cells||{}),
  added_rows:structuredClone(old.added_rows||[]),removed_rows:[...(old.removed_rows||[])]};
}
const pendingSaves=new Set();
let flushBatchInput=null;
function save(sheet,obj){
 const task=saveNow(sheet,obj);pendingSaves.add(task);
 task.then(()=>pendingSaves.delete(task),()=>pendingSaves.delete(task));return task;
}
async function flushPending(){
 if(flushBatchInput)await flushBatchInput();
 if(pendingSaves.size)await Promise.all([...pendingSaves]);
}
async function refreshPorkJelly(){
 if(!user())return;
 const id=approvedMushroomMasters['PK-J-INTERNAL'].master;
 const {data,error}=await client().from('jc_recipe_workbook_drafts').select('component_id,editor_id,cells,added_rows,removed_rows,updated_at').eq('component_id',id).eq('editor_id',user().id).maybeSingle();
 if(error)throw error;
 if(data){const old=A.drafts.get(id);if(!old?.updated_at||!data.updated_at||data.updated_at>=old.updated_at)A.drafts.set(id,data);}
 else if(!pendingSaves.size)A.drafts.delete(id);
 return data;
}
async function saveNow(sheet,obj){
 if(sheet.variantServing)throw Error('Fixed three-piece menu variation: edit the shared supporting recipe instead.');
 if(!canEdit())throw Error('Authorized editing session required');
 const payload={...obj,updated_at:new Date().toISOString()};
 sourceStatus('⏳ Saving changes automatically…');
 try{
 const group=matchingMasters(sheet),previous=state(sheet),payloads=[payload],seen=new Set([payload.component_id]);
 const changed=Object.keys(obj.cells||{}).filter(key=>JSON.stringify(obj.cells[key])!==JSON.stringify(previous.cells?.[key]));
 const structural=JSON.stringify(obj.added_rows||[])!==JSON.stringify(previous.added_rows||[])||JSON.stringify(obj.removed_rows||[])!==JSON.stringify(previous.removed_rows||[]);
 const recipeChange=structural||changed.some(key=>{const [row,col]=key.split(':').map(Number);return ingredientRow(sheet,row)||isYield(sheet,row,col);});
 for(const peer of recipeChange&&!mushroomMaster(sheet)?group.matches:[]){
  if(seen.has(peer.component_id))continue;seen.add(peer.component_id);
  const copy=currentDraft(peer);
  for(const key of changed){
   const [row,col]=key.split(':').map(Number);
   if(!ingredientRow(sheet,row)&&!isYield(sheet,row,col))continue;
   // Refuse to overwrite a separately edited conflicting ingredient or yield.
   if(Object.prototype.hasOwnProperty.call(copy.cells,key)&&JSON.stringify(copy.cells[key])!==JSON.stringify(previous.cells?.[key])&&JSON.stringify(copy.cells[key])!==JSON.stringify(obj.cells[key]))throw Error('Conflicting saved master draft for '+group.code+' in '+peer.file+'; review before synchronizing');
   copy.cells[key]=obj.cells[key];
  }
  if(structural){
   if(JSON.stringify(copy.added_rows||[])!==JSON.stringify(previous.added_rows||[])||JSON.stringify(copy.removed_rows||[])!==JSON.stringify(previous.removed_rows||[]))throw Error('Conflicting master ingredient rows in '+peer.file);
   copy.added_rows=structuredClone(obj.added_rows||[]);copy.removed_rows=structuredClone(obj.removed_rows||[]);
  }
  payloads.push({...copy,updated_at:payload.updated_at});
 }
 const {error}=await client().from('jc_recipe_workbook_drafts').upsert(payloads,{onConflict:'component_id,editor_id'});
 if(error)throw error;
 for(const p of payloads)A.drafts.set(p.component_id,p);
 if(payloads.some(p=>p.component_id===approvedMushroomMasters['PK-J-INTERNAL'].master)){
  const confirmed=await refreshPorkJelly();
  const expected=payloads.find(p=>p.component_id===approvedMushroomMasters['PK-J-INTERNAL'].master);
  if(!confirmed||String(confirmed.cells?.['4:6'])!==String(expected.cells?.['4:6']))throw Error('Pork Jelly saved yield could not be verified. Reload before editing.');
 }
 A.lastSync={id:sheet.id,ids:payloads.map(p=>p.component_id)};
 showDependencyReport(sheet,A.lastSync.ids);
 if(A.active===sheet)showLiveBreakdown(sheet);
 if(A.active&&A.active.id!==sheet.id&&!document.querySelector('#jeWorkbookGrid input:focus,#jeWorkbookGrid select:focus,#jeWorkbookGrid textarea:focus'))render(A.active,'jeWorkbookGrid');
 sourceStatus('✓ Saved '+payloads.length+' master draft'+(payloads.length===1?'':'s')+(group.code?' · '+group.code:'')+(group.conflicts.length?' · ⚠ Conflicting code copies require review':'')+' · Unpublished');
 }catch(error){sourceStatus('⚠ Changes not saved · '+(error.message||error));throw error;}
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
function isPorkJelly(sheet){return mushroomMaster(sheet)===approvedMushroomMasters['PK-J-INTERNAL'];}
function scaledBatchDraft(sheet,to){
 const src=canonicalSheet(sheet),address=yieldAddress(src),from=yieldValue(src);
 if(!(from>0&&to>0&&Number.isFinite(to)))throw Error('Enter a positive batch yield');
 const obj=currentDraft(src),factor=to/from;
 const removed=new Set((obj.removed_rows||[]).map(String));
 for(const row of src.rows||[]){
  const rr=Number(row.r);
  if(!ingredientRow(src,rr)||removed.has(String(rr)))continue;
  const q=cellValue(src,rr,0),product=productForRow(src,rr);
  if(q!==''&&q!=null&&product.code&&Number.isFinite(Number(q)))obj.cells[k(rr,0)]=Number((Number(q)*factor).toFixed(6));
 }
 for(const extra of obj.added_rows||[]){
  const q=extra.cells?.[0];if(q!==''&&q!=null&&Number.isFinite(Number(q)))extra.cells[0]=Number((Number(q)*factor).toFixed(6));
 }
 const unit=mushroomMaster(src)===approvedMushroomMasters['PK-J-INTERNAL']?'lb':String(cellValue(src,address.row,address.col)||src.yield||'').replace(/^\s*[\d,.]+\s*/,'').trim();
 obj.cells[k(address.row,address.col)]=unit?to+' '+unit:to;
 return obj;
}
async function setCell(sheet,row,col,value){
 if(isBlankNoteCell(sheet,row,col))throw Error('Notes and prep instructions are intentionally blank in the active Recipe Book.');
 if(!sheet.variantServing&&!isProtected(sheet)&&isYield(sheet,row,col)&&mushroomMaster(sheet)===approvedMushroomMasters['PK-J-INTERNAL']){
  const m=String(value).trim().match(/^(\d+(?:\.\d+)?)\s*(?:lb|lbs)?$/i);
  if(!m||!(Number(m[1])>0))throw Error('Enter the target Pork Jelly yield in pounds');
  const to=Number(m[1]);if(to===yieldValue(sheet))return true;
  await save(sheet,scaledBatchDraft(sheet,to));return true;
 }
 if(sheet.variantServing)throw Error('Fixed three-piece menu variation.');
 const registeredNote=String(original(sheet,row,col)).match(/\b(?:CT\s+)?internal\s+(?:product\s+)?code\s*:\s*([A-Z0-9_-]+)/i);
 if(registeredNote&&String(value).trim()!==''&&!String(value).toUpperCase().includes(registeredNote[1].toUpperCase()))throw Error('Use Update CT code to change the registered identifier. You may clear the display note without changing the registered code.');
 if((isProtected(sheet)||fillingReference(sheet,row))&&ingredientRow(sheet,row))throw Error('Main dish ingredients are calculated; edit the supporting recipe.');
 if(isProtected(sheet)&&isYield(sheet,row,col)){if(!/^\d+(?:\.\d+)?$/.test(String(value).trim())||!(Number(value)>0))throw Error('Enter a positive serving yield');value=Number(value);}
 if(isYield(sheet,row,col)&&dumplingCountYield(sheet,cellValue(sheet,row,col))){const m=String(value).match(/^\s*(\d+)\s*(?:dumplings?|pcs?|pieces?)?\s*$/i);if(!m||Number(m[1])<1)throw Error('Enter a positive whole-number dumpling yield');value=Number(m[1]);}
 var obj=currentDraft(sheet),key=k(row,col);
 const originalValue=cellValue(sheet,row,col);
 if(String(value)===String(originalValue))return true;
 if(col===0&&ingredientRow(sheet,row)&&value!==''&&(!Number.isFinite(Number(value))||Number(value)<0))throw Error('Ingredient quantity must be a nonnegative number');
 if(typeof row==='string'&&row.startsWith('new-')){
  const target=obj.added_rows.find(x=>x.id===row);
  if(!target)throw Error('Added ingredient row no longer exists');
  target.cells=target.cells||{};target.cells[col]=(col===0&&value!==''?Number(value)/preparedQuantityFactor(target.cells[3]):value);if(col===7&&value.includes(' | ')){const [code,...name]=value.split(' | ');target.cells[3]=code;target.cells[2]=name.join(' | ');}
 }else{
  obj.cells[key]=(col===0&&ingredientRow(sheet,row)&&value!==''?Number(value)/preparedQuantityFactor(productForRow(sheet,row).code):value);
  if(col===7&&String(value).includes(' | ')){
   const parts=String(value).split(' | '),code=parts.shift().trim(),productName=parts.join(' | ').trim();
   const actual=lookupProduct(code);
   if(actual&&normalizeName(actual.canonical_name)===normalizeName(productName)){obj.cells[k(row,2)]=actual.canonical_name;obj.cells[k(row,3)]=actual.ct_code;obj.cells[key]=actual.ct_code+' | '+actual.canonical_name;}
   else if(productName){obj.cells[k(row,2)]=productName;obj.cells[k(row,3)]=code;}
  }
  if(col===1&&ingredientRow(sheet,row)){
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
 if(isProtected(sheet))throw Error('Add ingredients to the supporting recipe, not the calculated main dish.');
 var max=Math.max(7,...(sheet.ingredients||[]).map(x=>Number(x.r)||7));
 var obj=currentDraft(sheet);
 obj.added_rows.push({id:'new-'+Date.now()+'-'+Math.random().toString(36).slice(2,9),after:max,cells:{0:'',1:'oz',2:'',3:'',4:'',5:'',6:'',7:''}});
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');
 return true;
}
async function removeRow(sheet,rowKey){
 if(isProtected(sheet))throw Error('Calculated main dish rows are protected.');
 if(rowKey==null||!ingredientRow(sheet,rowKey))throw Error('First select an ingredient row to remove');
 if(!confirm('Remove this worksheet row from the working draft? Your original Excel workbook is protected.'))return false;
 var obj=currentDraft(sheet);
 if(String(rowKey).startsWith('new-'))obj.added_rows=obj.added_rows.filter(x=>x.id!==rowKey);
 else if(!obj.removed_rows.map(String).includes(String(rowKey)))obj.removed_rows.push(Number(rowKey));
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');A.selectedRow=null;return true;
}
async function restoreRows(sheet){
 if(isProtected(sheet))throw Error('Calculated main dish rows are protected.');
 var obj=currentDraft(sheet);obj.removed_rows=[];
 await save(sheet,obj);render(sheet,'jeWorkbookGrid');
}
async function publish(sheet){
 sheet=canonicalSheet(sheet);
 if(sheet.variantServing)throw Error('Publish the shared source recipe instead.');
 if(bridge()?.getRole()!=='admin')throw Error('Owner approval is required');
 if(!A.drafts.has(sheet.component_id))throw Error('No unpublished changes to publish');
 const count=(bridge()?.getLibrary()?.sheets||[]).filter(x=>mushroomMaster(sheet)?mushroomMaster(sheet).ids.includes(x.component_id):x.component_id===sheet.component_id).length;
 if(!confirm('Publish this master component to '+count+' matching worksheet'+(count!==1?'s':'')+'? This will approve shared changes.'))return;
 const {data,error}=await client().rpc('jc_recipe_publish_workbook_component',{p_component_id:sheet.component_id});
 if(error)throw error;
 await load();sourceStatus('Published master component · version '+data);
 render(sheet,'jeWorkbookGrid');
}
async function scaleBatch(sheet){
 if(sheet.variantServing)throw Error('Secret Menu serving is fixed at three pieces.');
 if(!canEdit())throw Error('Authorized editor required');
 const address=yieldAddress(sheet);const from=yieldValue(sheet)||0;
 if(!(from>0))throw Error('A numeric source target yield is required before scaling');
 const unit=String(cellValue(sheet,address.row,address.col)||sheet.yield||'').replace(/^\s*[\d,.]+\s*/,'').trim();
 const result=prompt('New target batch yield'+(unit?' ('+unit+')':'')+':',fmt(from));
 if(result===null)return;
 const to=Number(String(result).trim().replace(/\s*lbs?$/i,''));
 if(!(to>0)||!Number.isFinite(to))throw Error('Enter a positive target yield');
 if(isProtected(sheet)){await setCell(sheet,address.row,address.col,to);render(sheet,'jeWorkbookGrid');return;}
 const factor=to/from;
 if(!confirm('Scale '+sheet.name+' by '+factor.toFixed(3)+' ×? Ingredient quantities change proportionally. Actual finished yield must still be measured.'))return;
 await save(sheet,scaledBatchDraft(sheet,to));render(sheet,'jeWorkbookGrid');
 sourceStatus('✓ Saved · target yield and ingredient quantities scaled together');
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
 let old=String(col===7&&!String(row).startsWith('new-')?productForRow(sheet,row).selector:row.toString().startsWith('new-')?(state(sheet).added_rows||[]).find(x=>x.id===row)?.cells?.[col]||'':cellValue(sheet,row,col));
 if(col===0&&ingredientRow(sheet,row)&&old!==''){const code=String(row).startsWith('new-')?(state(sheet).added_rows||[]).find(x=>x.id===row)?.cells?.[3]:productForRow(sheet,row).code;old=String(linkedPreparedQuantity(old,code));}
 var control;
 if(col===7&&ingredientRow(sheet,row)){
  control=document.createElement('select');
  const sourceCode=String(original(sheet,row,3)||''),sourceName=String(original(sheet,row,2)||''); const sourceChoice=sourceCode&&sourceName?sourceCode+' | '+sourceName:'';
  const values=[...new Set(['',old,sourceChoice,...A.products.map(p=>p.ct_code+' | '+p.canonical_name)])];
  values.forEach(v=>{const opt=document.createElement('option');opt.value=v;opt.textContent=v||'Select a product…';if(v===sourceChoice&&v!==old)opt.textContent+=' (Original recipe)';if(!v)opt.disabled=true;control.appendChild(opt)});control.value=old;
 }else if(col===1&&ingredientRow(sheet,row)){
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
 const initialValue=control.value;
 let committed=false;
 async function complete(commit=true){
  if(committed)return;committed=true;
  try{if(commit&&control.value!==initialValue){if(col===7&&ingredientRow(sheet,row)&&old&&control.value!==old&&!confirm('Replace '+old+' with '+control.value+'?')){render(sheet,'jeWorkbookGrid');return;}await setCell(sheet,row,col,control.value)}}catch(e){sourceStatus('Not saved: '+e.message)}
  if(A.active===sheet){
 if(isYield(sheet,row,col)||!document.querySelector('#jeWorkbookGrid .jc-wb-edit-control:focus'))render(sheet,'jeWorkbookGrid');
 else {td.replaceChildren();const span=document.createElement('span');span.className='jc-wb-value';span.textContent=rendered(sheet,row,col,{v:original(sheet,row,col)});td.appendChild(span);}
 }
 }
 control.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();complete()}if(e.key==='Escape'){e.preventDefault();complete(false)}});
 if(isProtected(sheet)&&isYield(sheet,row,col)){
 control.addEventListener('input',()=>{
 const n=Number(control.value),base=yieldValue(sheet,true);if(!(n>0&&base>0))return;
 for(const node of document.querySelectorAll('#jeWorkbookGrid td[data-c="0"] .jc-wb-value')){
 const rr=Number(node.closest('td').dataset.r);if(!ingredientRow(sheet,rr))continue;
 const sourceCell=(sheet.rows||[]).find(x=>Number(x.r)===rr)?.c?.find(x=>Number(x.c)===0);
 const q=linkedFormulaValue(sheet,rr,0,sourceCell)??original(sheet,rr,0);
 if(q!==''&&q!=null&&Number.isFinite(Number(q)))node.textContent=fmt(Number(q)*n/base);
 }
 });
 }
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
 A.selectedCell=picker||null;
 if(picker&&A.active&&!picker.querySelector('input,select,textarea'))editCell(picker);
});
document.addEventListener('keydown',async e=>{
 if(!['Delete','Backspace'].includes(e.key)||e.target.closest?.('input,textarea,select,[contenteditable]'))return;
 const td=A.selectedCell,sheet=A.active;if(!td?.isConnected||!sheet||!canEdit())return;
 const row=Number(td.dataset.r),col=Number(td.dataset.c);
 if(!Number.isFinite(row)||ingredientRow(sheet,row)||row<=notesRow(sheet)||col!==1)return;
 if(/Prep\s*\/\s*Plating Instructions/i.test(String(original(sheet,row,col))))return;
 e.preventDefault();try{await setCell(sheet,row,col,'');if(A.active===sheet)render(sheet,'jeWorkbookGrid');}catch(err){sourceStatus('Not saved: '+err.message);}
});
window.JCRecipeAdvanced={
 setProtectedMain:id=>{protectedMain=id||null;},
 load,render,addRow,removeRow,restoreRows,publish,changeCode,scaleBatch,print,groupConvert,flushPending,
 build:'20261124-breakdown19',refreshPorkJelly,
 get selectedRow(){return A.selectedRow},
 get active(){return A.active},
 get productCount(){return A.products.length},
 get version(){return A.published.get(A.active?.component_id)?.version||0}
};
})();