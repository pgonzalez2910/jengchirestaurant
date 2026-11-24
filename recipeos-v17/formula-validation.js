/* RecipeOS V17 shadow validation. No database writes. Engine supplied by caller. */
(function(root){
'use strict';
const own=(o,k)=>Object.prototype.hasOwnProperty.call(o||{},k);
const norm=s=>String(s??'').trim().toLowerCase();
function cellAddress(row,col){return {row:Number(row)-1,col:Number(col)};}
function buildInput(sheets,overrides=new Map()){
 const data=Object.create(null);
 for(const sheet of sheets){
  const name=String(sheet.sheet||sheet.name);
  if(own(data,name))throw Error('Duplicate worksheet name: '+name);
  const arr=[];data[name]=arr;
  const edits=overrides.get(sheet.component_id)||{};
  for(const row of sheet.rows||[]){
   const n=Number(row.r)-1;if(!Number.isInteger(n)||n<0)throw Error('Invalid row address');
   arr[n]??=[];
   for(const cell of row.c||[]){
    const col=Number(cell.c),key=row.r+':'+col;
    if(!Number.isInteger(col)||col<0)throw Error('Invalid column address');
    arr[n][col]=own(edits.cells,key)?edits.cells[key]:cell.f?String(cell.f).startsWith('=')?cell.f:'='+cell.f:cell.v??null;
   }
  }
  for(const [key,value] of Object.entries(edits.cells||{})){
   const [r,c]=key.split(':').map(Number);
   if(!Number.isInteger(r)||r<1||!Number.isInteger(c)||c<0)continue;
   arr[r-1]??=[];arr[r-1][c]=value;
  }
  // Structural edits need explicit address mapping, not silent shifting of archived references.
  if(edits.added_rows?.length||edits.removed_rows?.length)throw Error('Structural draft requires reference mapping: '+name);
  for(let i=0;i<arr.length;i++)arr[i]=Array.from(arr[i]||[],v=>v??null);
 }
 return data;
}
function equal(a,b){
 if(a==null||a==='')return b==null||b==='';
 if(typeof a==='number'&&typeof b==='number')return Math.abs(a-b)<=1e-8*Math.max(1,Math.abs(a),Math.abs(b));
 return a===b;
}
function auditWorkbook(engineFactory,sheets,overrides=new Map()){
 if(typeof engineFactory!=='function')throw Error('Supply an approved formula engine factory');
 const engine=engineFactory(buildInput(sheets,overrides));
 const result={formulas:0,matched:0,mismatched:0,errors:0,unverified:0,issues:[]};
 try{
  for(const sheet of sheets){
   const name=String(sheet.sheet||sheet.name),id=engine.getSheetId(name);
   for(const row of sheet.rows||[])for(const cell of row.c||[]){
    if(!cell.f)continue;result.formulas++;
    const key=row.r+':'+cell.c;
    if(own(overrides.get(sheet.component_id)?.cells,key)){result.unverified++;continue;}
    const address={sheet:id,...cellAddress(row.r,cell.c)};
    let value;
    try{value=engine.getCellValue(address);}catch(e){value={error:String(e.message||e)};}
    const err=value&&typeof value==='object';
    if(err)result.errors++;
    else if(!own(cell,'v'))result.unverified++;
    else if(equal(cell.v,value))result.matched++;
    else result.mismatched++;
    if(err||own(cell,'v')&&!equal(cell.v,value))result.issues.push({worksheet:name,row:row.r,col:cell.c,formula:cell.f,expected:cell.v,actual:err?String(value.message||value.error||value.type||'Calculation error'):value});
   }
  }
 }finally{engine.destroy?.();}
 return result;
}
function auditLibrary(engineFactory,sheets){
 const groups=new Map();
 for(const sheet of sheets){if(!groups.has(sheet.file))groups.set(sheet.file,[]);groups.get(sheet.file).push(sheet);}
 const results=[];
 for(const [file,group] of groups){try{results.push({file,...auditWorkbook(engineFactory,group)});}catch(e){results.push({file,blocked:String(e.message||e)});}}
 return results;
}
function testServing(engineFactory){
 const sheets=[{sheet:'Filling',component_id:'f',rows:[{r:3,c:[{c:0,v:1860}]},{r:8,c:[{c:0,v:589.3069306930693}]}]},{sheet:'Main Dish',component_id:'m',rows:[{r:3,c:[{c:0,v:6}]},{r:8,c:[{c:0,f:"='Filling'!A8*A3/'Filling'!A3",v:589.3069306930693*6/1860}]}]}];
 const engine=engineFactory(buildInput(sheets));
 try{
  const main=engine.getSheetId('Main Dish'),fill=engine.getSheetId('Filling'),at={sheet:main,row:7,col:0};
  const initial=engine.getCellValue(at);
  engine.setCellContents({sheet:main,row:2,col:0},[[8]]);
  const eight=engine.getCellValue(at);
  engine.setCellContents({sheet:fill,row:7,col:0},[[1178.6138613861386]]);
  const doubled=engine.getCellValue(at);
  return {initial,eight,doubled,passed:equal(eight,initial*8/6)&&equal(doubled,eight*2)};
 }finally{engine.destroy?.();}
}
const api={buildInput,auditWorkbook,auditLibrary,testServing};
if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.JCFormulaValidation=api;
})(typeof globalThis!=='undefined'?globalThis:this);
