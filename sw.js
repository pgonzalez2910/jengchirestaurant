const CACHE='jc-place-order-shell-v1';
const IMAGE_CACHE='jc-place-order-images-v1';
const DB_NAME='jc-place-order-offline-v1', DB_VERSION=1;
const SUPABASE_URL="https://kpldzwlftkvjjntgsqxx.supabase.co";
const SUPABASE_ANON_KEY="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtwbGR6d2xmdGt2ampudGdzcXh4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDY2MzE5NzAsImV4cCI6MjA2MjIwNzk3MH0.qnWbOQv2RLPsIyO-oRwQkAN2VhmmdhTBt46SweUsLbs";
self.addEventListener('install',event=>event.waitUntil((async()=>{const c=await caches.open(CACHE);await c.addAll(['./','./jengchi-place-order.html']);const external=['https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2','https://cdn.jsdelivr.net/npm/sortablejs@1.15.6/Sortable.min.js'];await Promise.allSettled(external.map(url=>c.add(url)));await self.skipWaiting()})()));
self.addEventListener('activate',event=>event.waitUntil(self.clients.claim()));
self.addEventListener('fetch',event=>{if(event.request.method!=='GET')return;const u=new URL(event.request.url);if(u.hostname==='raw.githubusercontent.com'||u.hostname==='cdn.jsdelivr.net'){event.respondWith(caches.open(IMAGE_CACHE).then(async c=>{const hit=await c.match(event.request);if(hit)return hit;try{const r=await fetch(event.request);if(r.ok||r.type==='opaque')c.put(event.request,r.clone());return r}catch(e){return hit||Response.error()}}));return}if(u.origin===self.location.origin){event.respondWith(fetch(event.request).then(r=>{const x=r.clone();caches.open(CACHE).then(c=>c.put(event.request,x));return r}).catch(()=>caches.match(event.request)))}});
function openDb(){return new Promise((resolve,reject)=>{const r=indexedDB.open(DB_NAME,DB_VERSION);r.onupgradeneeded=()=>{const d=r.result;if(!d.objectStoreNames.contains('drafts'))d.createObjectStore('drafts',{keyPath:'key'});if(!d.objectStoreNames.contains('submissions'))d.createObjectStore('submissions',{keyPath:'client_submission_id'});if(!d.objectStoreNames.contains('meta'))d.createObjectStore('meta',{keyPath:'key'})};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
async function getAll(store){const d=await openDb();return new Promise((res,rej)=>{const tx=d.transaction(store,'readonly');const r=tx.objectStore(store).getAll();r.onsuccess=()=>res(r.result||[]);r.onerror=()=>rej(r.error)})}
async function put(store,val){const d=await openDb();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).put(val);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
async function rest(path,method,body,prefer=''){const h={'apikey':SUPABASE_ANON_KEY,'Authorization':'Bearer '+SUPABASE_ANON_KEY,'Content-Type':'application/json'};if(prefer)h['Prefer']=prefer;const r=await fetch(SUPABASE_URL+'/rest/v1/'+path,{method,headers:h,body:body?JSON.stringify(body):undefined});if(!r.ok)throw new Error(await r.text());return r}
async function syncDraft(d){const payload={order_date:d.order_date,user_type:d.user_type,product_name:d.product_name,vendor:d.vendor||null,qty:Number(d.qty||0),item_code:d.item_code||null,image_url:d.image_url||null,code_source:d.code_source||null,client_updated_at:d.updated_at,updated_at:new Date().toISOString()};await rest('place_order_drafts?on_conflict=order_date,user_type,product_name','POST',[payload],'resolution=merge-duplicates');const compat={order_date:d.order_date,product_name:d.product_name,user:d.user_type==='rony'?'Rony':'Julio',tentative:Number(d.qty||0),vendor:d.vendor||null,item_code:d.item_code||null,image_url:d.image_url||null};try{await rest('inventory_check?on_conflict=order_date,product_name,user','POST',[compat],'resolution=merge-duplicates')}catch(e){}d.dirty=false;d.server_synced_at=new Date().toISOString();await put('drafts',d)}
async function sendSubmission(s){for(const i of s.items){const d=(await getAll('drafts')).find(x=>x.order_date===s.order_date&&x.user_type===s.user_type&&x.product_name===i.product_name);if(d&&d.dirty)await syncDraft(d)}try{await rest('place_order_submissions?on_conflict=client_submission_id','POST',[{client_submission_id:s.client_submission_id,order_date:s.order_date,user_type:s.user_type,submitted_at:s.submitted_at,item_count:s.items.length,payload:{items:s.items},email_status:'pending'}],'resolution=ignore-duplicates')}catch(e){}const lines=s.items.map(i=>({order_date:s.order_date,submitted_at:s.submitted_at,submitted_by:s.user_type,product_name:i.product_name,quantity:i.quantity,vendor:i.vendor,item_code:i.item_code,user_section:s.user_type}));try{await rest('submitted_orders?on_conflict=order_date,product_name,user_section','POST',lines,'resolution=merge-duplicates')}catch(e){}const r=await fetch(SUPABASE_URL+'/functions/v1/send-order-notification',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+SUPABASE_ANON_KEY},body:JSON.stringify({submissionId:s.client_submission_id,submittedBy:s.user_type,orderDate:s.order_date,companyLogo:'https://raw.githubusercontent.com/pgonzalez2910/jengchi-product-images/main/jengchilogo.jpg',items:s.items})});if(!r.ok)throw new Error(await r.text());const x=await r.json();if(!x.success)throw new Error(x.error||'Email failed');s.status='synced';s.email_status='sent';s.synced_at=new Date().toISOString();await put('submissions',s)}
async function syncPendingCatalogOrders(){
  const meta=await getAll('meta');
  const jobs=meta.filter(x=>String(x.key||'').startsWith('catalog-order-pending:')&&Array.isArray(x.value)&&x.value.length);
  for(const job of jobs){
    let complete=true;
    for(const row of job.value){
      try{
        await rest(`place_order_catalog?id=eq.${encodeURIComponent(row.id)}`,'PATCH',{
          sort_order:Number(row.sort_order),
          updated_at:new Date().toISOString()
        });
      }catch(e){
        complete=false;
        break;
      }
    }
    if(complete)await put('meta',{key:job.key,value:[]});
  }
}
async function backgroundSync(){await syncPendingCatalogOrders();const drafts=(await getAll('drafts')).filter(x=>x.dirty);for(const d of drafts){try{await syncDraft(d)}catch(e){}}const subs=(await getAll('submissions')).filter(x=>x.status!=='synced'||x.email_status!=='sent');for(const s of subs){try{await sendSubmission(s)}catch(e){}}const clients=await self.clients.matchAll({type:'window'});clients.forEach(c=>c.postMessage({type:'SYNC_NOW'}))}
self.addEventListener('sync',event=>{if(event.tag==='jc-place-order-sync')event.waitUntil(backgroundSync())}); 
