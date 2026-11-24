/* Original dependency-free RecipeOS formula subset. Unsupported expressions fail explicitly. */
(function(root){
'use strict';
function factory(data){
 const names=Object.keys(data), sheets=names.map(n=>data[n]), ids=new Map(names.map((n,i)=>[n.toLowerCase(),i]));
 const err=m=>{throw Error(m)};
 const num=x=>{if(x===null||x===''||x===undefined)return 0;const n=Number(x);return Number.isFinite(n)?n:err('VALUE')};
 const flat=a=>a.flat(Infinity);
 const equal=(a,b)=>typeof a==='string'&&typeof b==='string'?a.toLowerCase()===b.toLowerCase():a===b;
 const col=s=>[...s.toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;
 function parse(f){
 const re=/\s*(?:"((?:[^"]|"")*)"|'((?:[^']|'')+)'!|([A-Za-z_][A-Za-z0-9_. ]*)!|(\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}(?=\s*:)|\$[A-Za-z]{1,3})(?![A-Za-z0-9_])|(\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|\.\d+)|([A-Za-z_][A-Za-z0-9_.]*)|(<>|<=|>=|[+\-*/^&=<>%,():]))/gy;
 let i=0,t=[];f=f.replace(/^=/,'').trim();
 while(i<f.length){re.lastIndex=i;const m=re.exec(f);if(!m)err('Unsupported token at '+f.slice(i));i=re.lastIndex;t.push(m[1]!==undefined?['str',m[1].replace(/""/g,'"')]:m[2]!==undefined?['sheet',m[2].replace(/''/g,"'")]:m[3]!==undefined?['sheet',m[3]]:m[4]?['ref',m[4]]:m[5]?['num',Number(m[5])]:m[6]?['id',m[6].toUpperCase()]:[m[7],m[7]]);}
 let p=0;const peek=()=>t[p]?.[0],take=()=>t[p++],expect=k=>{if(peek()!==k)err('Expected '+k);return take()};
 function primary(){let x=take();if(!x)err('Missing operand');let [k,v]=x;
 if(k==='num'||k==='str')return [k,v];
 if(k==='+'||k==='-')return ['un',k,primary()];
 if(k==='('){let a=expr(0);expect(')');return a;}
 if(k==='sheet'||k==='ref'){let s=k==='sheet'?v:null,r=k==='sheet'?expect('ref')[1]:v;if(peek()===':'){take();const end=take();if(!end||!['ref','id'].includes(end[0]))err('Expected range end');return ['range',s,r,end[1]];}return ['ref',s,r];}
 if(k==='id'){if(v==='TRUE'||v==='FALSE')return ['num',v==='TRUE'];expect('(');let a=[];if(peek()!==')'){do{a.push(expr(0));if(peek()!==',')break;take();}while(true);}expect(')');return ['call',v,a];}
 err('Unsupported operand '+k);}
 const prec={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
 function expr(min){let a=primary();while(peek()==='%'){take();a=['bin','/',a,['num',100]];}while((prec[peek()]||0)>min){let op=take()[0];a=['bin',op,a,expr(prec[op]-(op==='^'?1:0))];}return a;}
 const ast=expr(0);if(p!==t.length)err('Trailing expression');return ast;
 }
 const cache=new Map(),active=new Set();
 function get(sheet,row,c){const key=[sheet,row,c].join(':');if(cache.has(key))return cache.get(key);if(active.has(key))err('CIRCULAR');active.add(key);try{const v=sheets[sheet]?.[row]?.[c]??null;const out=typeof v==='string'&&v.startsWith('=')?evaluate(parse(v),sheet,row):v;cache.set(key,out);return out;}finally{active.delete(key);}}
 function address(r){const m=r.replace(/\$/g,'').match(/^([A-Z]+)(\d+)$/i);return [Number(m[2])-1,col(m[1])];}
 function evaluate(a,s,row){const [k,x,y,z]=a;const ev=b=>evaluate(b,s,row);
 if(k==='num'||k==='str')return x;
 if(k==='ref'||k==='range'){const sh=x===null?s:ids.get(x.toLowerCase());if(sh===undefined)err('REF '+x);const whole=k==='range'&&/^\$?[A-Z]{1,3}$/i.test(y)&&/^\$?[A-Z]{1,3}$/i.test(z);const [r,c]=whole?[0,col(y.replace(/\$/g,''))]:address(y);if(k==='ref')return get(sh,r,c);const [r2,c2]=whole?[Math.max(0,sheets[sh].length-1),col(z.replace(/\$/g,''))]:address(z);if((r2-r+1)*(c2-c+1)>100000)err('Range limit');return Array.from({length:r2-r+1},(_,i)=>Array.from({length:c2-c+1},(_,j)=>get(sh,r+i,c+j)));}
 if(k==='un')return (x==='-'?-1:1)*num(ev(y));
 if(k==='bin'){let l=ev(y),r=ev(z);switch(x){case '=':return equal(l??'',r??'');case '<>':return !equal(l??'',r??'');case '<':return l<r;case '>':return l>r;case '<=':return l<=r;case '>=':return l>=r;case '&':return String(l??'')+String(r??'');case '+':return num(l)+num(r);case '-':return num(l)-num(r);case '*':return num(l)*num(r);case '/':if(num(r)===0)err('DIV/0');return num(l)/num(r);case '^':return num(l)**num(r);}}
 if(k==='call'){if(x==='IF')return ev(y[0])?ev(y[1]):y[2]?ev(y[2]):false;if(x==='IFERROR'){try{return ev(y[0]);}catch(e){if(/Unsupported|Trailing|Expected|Missing operand|Range limit/i.test(e.message))throw e;return ev(y[1]);}}
 const v=y.map(ev),a0=v[0],a1=v[1],a2=v[2],str=q=>String(q??'');
 switch(x){case 'SUM':return flat(v).reduce((n,q)=>n+(typeof q==='number'?q:0),0);case 'COUNT':return flat(v).filter(q=>typeof q==='number').length;case 'MAX':case 'MIN':{const ns=flat(v).filter(q=>typeof q==='number');return ns.length?(x==='MAX'?Math.max(...ns):Math.min(...ns)):0;}case 'AVERAGE':{let n=flat(v).filter(q=>typeof q==='number');return n.length?n.reduce((a,b)=>a+b,0)/n.length:err('DIV/0');}case 'AND':return flat(v).every(Boolean);case 'OR':return flat(v).some(Boolean);case 'NOT':return !a0;case 'ISNUMBER':return typeof a0==='number';case 'LEN':return str(a0).length;case 'LOWER':return str(a0).toLowerCase();case 'TRIM':return str(a0).trim().replace(/ +/g,' ');case 'LEFT':return str(a0).slice(0,a1??1);case 'RIGHT':return str(a0).slice(-(a1??1));case 'MID':return str(a0).substr(num(a1)-1,num(a2));case 'FIND':case 'SEARCH':{let needle=str(a0),hay=str(a1);if(x==='SEARCH'){needle=needle.toLowerCase();hay=hay.toLowerCase();}let i=hay.indexOf(needle,(v[2]??1)-1);return i<0?err('VALUE'):i+1;}case 'VALUE':return num(a0);case 'INT':return Math.floor(num(a0));case 'ROUND':case 'ROUNDUP':case 'ROUNDDOWN':{let p=10**num(a1),n=num(a0)*p;return Math.sign(n)*(x==='ROUND'?Math.round(Math.abs(n)):x==='ROUNDUP'?Math.ceil(Math.abs(n)):Math.floor(Math.abs(n)))/p;}case 'ROW':return y.length&&y[0][0]==='ref'?address(y[0][2])[0]+1:row+1;case 'MATCH':{if(a2!==0)err('Unsupported approximate MATCH');let i=flat([a1]).findIndex(q=>equal(q,a0));return i<0?err('N/A'):i+1;}case 'INDEX':return a0?.[num(a1)-1]?.[num(a2??1)-1]??null;case 'VLOOKUP':{if(a2===undefined||v[3]!==false&&v[3]!==0)err('Unsupported approximate VLOOKUP');let r=a1.find(q=>equal(q[0],a0));return r?r[num(a2)-1]:err('N/A');}case 'SUBSTITUTE':{const s=str(a0),old=str(a1),replacement=str(a2);if(!old)return s;if(v[3]==null)return s.split(old).join(replacement);let occurrence=0;return s.split(old).map((part,i)=>i===0?part:(++occurrence===num(v[3])?replacement:old)+part).join('');}case 'TEXT':{const pattern=str(a1);if(/^0(?:\.0+)?$/.test(pattern))return num(a0).toFixed((pattern.split('.')[1]||'').length);if(pattern==='0%')return (num(a0)*100).toFixed(0)+'%';err('Unsupported TEXT format '+pattern);}case 'HYPERLINK':return a1??a0;default:err('Unsupported function '+x);}}
 err('Unsupported AST');}
 return {getSheetId:n=>ids.get(n.toLowerCase()),getCellValue:a=>{try{return get(a.sheet,a.row,a.col)}catch(e){return {error:e.message}}},setCellContents(a,rows){for(let i=0;i<rows.length;i++){sheets[a.sheet][a.row+i]??=[];for(let j=0;j<rows[i].length;j++)sheets[a.sheet][a.row+i][a.col+j]=rows[i][j];}cache.clear();},destroy(){cache.clear();}};
}
if(typeof module!=='undefined'&&module.exports)module.exports=factory;else root.JCFreeFormulaEngine=factory;
})(typeof globalThis!=='undefined'?globalThis:this);
