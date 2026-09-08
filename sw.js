const CACHE='jc-place-order-shell-v26';
const IMAGE_CACHE='jc-place-order-images-v26';
const DB_NAME='jc-place-order-offline-v1', DB_VERSION=1;

const SUPABASE_URL="https://kpldzwlftkvjjntgsqxx.supabase.co";
const SUPABASE_ANON_KEY="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtwbGR6d2xmdGt2ampudGdzcXh4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDY2MzE5NzAsImV4cCI6MjA2MjIwNzk3MH0.qnWbOQv2RLPsIyO-oRwQkAN2VhmmdhTBt46SweUsLbs";

const PLACE_ORDER_PAGE='./jengchiorder.html';
const LEGACY_PLACE_ORDER_PAGE='./jengchi-place-order.html';
const INVENTORY_PAGE='./JengChi_Inventory_Count_Sheet_v8.html';

self.addEventListener('install',event=>event.waitUntil((async()=>{
  const c=await caches.open(CACHE);

  await Promise.allSettled([
    c.add('./'),
    c.add(PLACE_ORDER_PAGE),
    c.add(INVENTORY_PAGE)
  ]);

  const external=[
    'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2',
    'https://cdn.jsdelivr.net/npm/sortablejs@1.15.6/Sortable.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css'
  ];

  await Promise.allSettled(external.map(url=>c.add(url)));
  await self.skipWaiting();
})()));

self.addEventListener('activate',event=>event.waitUntil((async()=>{
  const keys=await caches.keys();

  await Promise.all(keys.map(k=>{
    if(
      (k.startsWith('jc-place-order-shell-')||k.startsWith('jc-place-order-images-')) &&
      k!==CACHE &&
      k!==IMAGE_CACHE
    ){
      return caches.delete(k);
    }
  }));

  await self.clients.claim();
})()));

self.addEventListener('fetch',event=>{
  if(event.request.method!=='GET')return;

  const u=new URL(event.request.url);

  if(u.hostname.endsWith('.supabase.co'))return;

  const isPlaceOrder=
    u.pathname.endsWith('/jengchiorder.html') ||
    u.pathname.endsWith('/jengchi-place-order.html');

  const isInventory=
    u.pathname.endsWith('/JengChi_Inventory_Count_Sheet_v8.html');

  // Network-first for HTML/navigation so GitHub/hosting updates appear immediately.
  if(event.request.mode==='navigate' || isPlaceOrder || isInventory){
    event.respondWith((async()=>{
      try{
        const fresh=await fetch(event.request,{
          cache:'no-store',
          headers:{'Cache-Control':'no-cache'}
        });

        if(fresh && fresh.ok){
          const c=await caches.open(CACHE);
          await c.put(event.request,fresh.clone());

          if(isPlaceOrder){
            await c.put(PLACE_ORDER_PAGE,fresh.clone()).catch(()=>{});
          }
        }

        return fresh;
      }catch(e){
        const exact=await caches.match(event.request);
        if(exact)return exact;

        if(isInventory){
          return (await caches.match(INVENTORY_PAGE)) || Response.error();
        }

        if(isPlaceOrder || event.request.mode==='navigate'){
          return (await caches.match(PLACE_ORDER_PAGE))
            || (await caches.match(LEGACY_PLACE_ORDER_PAGE))
            || (await caches.match('./'))
            || Response.error();
        }

        return Response.error();
      }
    })());

    return;
  }

  if(
    u.hostname==='raw.githubusercontent.com' ||
    u.hostname==='cdn.jsdelivr.net' ||
    u.hostname==='cdnjs.cloudflare.com'
  ){
    event.respondWith((async()=>{
      const c=await caches.open(IMAGE_CACHE);
      const hit=await c.match(event.request);

      try{
        const r=await fetch(event.request,{cache:'no-store'});

        if(r.ok||r.type==='opaque'){
          await c.put(event.request,r.clone());
        }

        return r;
      }catch(e){
        return hit||Response.error();
      }
    })());

    return;
  }

  if(u.origin===self.location.origin){
    event.respondWith(
      fetch(event.request,{cache:'no-store'})
        .then(async r=>{
          if(r && r.ok){
            const c=await caches.open(CACHE);
            c.put(event.request,r.clone()).catch(()=>{});
          }
          return r;
        })
        .catch(()=>caches.match(event.request))
    );
  }
});

self.addEventListener('message',event=>{
  const data=event.data||{};

  if(data.type==='SKIP_WAITING'){
    self.skipWaiting();
    return;
  }

  if(data.type!=='CACHE_URL'||!data.url)return;

  event.waitUntil((async()=>{
    try{
      const c=await caches.open(CACHE);
      const r=await fetch(data.url,{cache:'no-store'});

      if(r && r.ok){
        await c.put(data.url,r.clone());
      }
    }catch(e){}
  })());
});

function openDb(){
  return new Promise((resolve,reject)=>{
    const r=indexedDB.open(DB_NAME,DB_VERSION);

    r.onupgradeneeded=()=>{
      const d=r.result;

      if(!d.objectStoreNames.contains('drafts')){
        d.createObjectStore('drafts',{keyPath:'key'});
      }

      if(!d.objectStoreNames.contains('submissions')){
        d.createObjectStore('submissions',{keyPath:'client_submission_id'});
      }

      if(!d.objectStoreNames.contains('meta')){
        d.createObjectStore('meta',{keyPath:'key'});
      }
    };

    r.onsuccess=()=>resolve(r.result);
    r.onerror=()=>reject(r.error);
  });
}

async function getAll(store){
  const d=await openDb();

  return new Promise((res,rej)=>{
    const tx=d.transaction(store,'readonly');
    const r=tx.objectStore(store).getAll();

    r.onsuccess=()=>res(r.result||[]);
    r.onerror=()=>rej(r.error);
  });
}

async function put(store,val){
  const d=await openDb();

  return new Promise((res,rej)=>{
    const tx=d.transaction(store,'readwrite');
    tx.objectStore(store).put(val);

    tx.oncomplete=()=>res();
    tx.onerror=()=>rej(tx.error);
  });
}

async function rest(path,method,body,prefer=''){
  const h={
    'apikey':SUPABASE_ANON_KEY,
    'Authorization':'Bearer '+SUPABASE_ANON_KEY,
    'Content-Type':'application/json'
  };

  if(prefer)h['Prefer']=prefer;

  const r=await fetch(
    SUPABASE_URL+'/rest/v1/'+path,
    {
      method,
      headers:h,
      body:body?JSON.stringify(body):undefined
    }
  );

  if(!r.ok)throw new Error(await r.text());

  return r;
}

async function syncDraft(d){
  const payload={
    order_date:d.order_date,
    user_type:d.user_type,
    product_name:d.product_name,
    vendor:d.vendor||null,
    qty:Number(d.qty||0),
    item_code:d.item_code||null,
    image_url:d.image_url||null,
    code_source:d.code_source||null,
    client_updated_at:d.updated_at,
    updated_at:new Date().toISOString()
  };

  await rest(
    'place_order_drafts?on_conflict=order_date,user_type,product_name',
    'POST',
    [payload],
    'resolution=merge-duplicates'
  );

  const compat={
    order_date:d.order_date,
    product_name:d.product_name,
    user:d.user_type==='rony'?'Rony':'Julio',
    tentative:Number(d.qty||0),
    vendor:d.vendor||null,
    item_code:d.item_code||null,
    image_url:d.image_url||null
  };

  try{
    await rest(
      'inventory_check?on_conflict=order_date,product_name,user',
      'POST',
      [compat],
      'resolution=merge-duplicates'
    );
  }catch(e){}

  d.dirty=false;
  d.server_synced_at=new Date().toISOString();

  await put('drafts',d);
}

async function sendSubmission(s){
  for(const i of s.items){
    const d=(await getAll('drafts')).find(
      x=>
        x.order_date===s.order_date &&
        x.user_type===s.user_type &&
        x.product_name===i.product_name
    );

    if(d&&d.dirty){
      await syncDraft(d);
    }
  }

  try{
    await rest(
      'place_order_submissions?on_conflict=client_submission_id',
      'POST',
      [{
        client_submission_id:s.client_submission_id,
        order_date:s.order_date,
        user_type:s.user_type,
        submitted_at:s.submitted_at,
        item_count:s.items.length,
        payload:{items:s.items},
        email_status:'pending'
      }],
      'resolution=ignore-duplicates'
    );
  }catch(e){}

  const lines=s.items.map(i=>({
    order_date:s.order_date,
    submitted_at:s.submitted_at,
    submitted_by:s.user_type,
    product_name:i.product_name,
    quantity:i.quantity,
    vendor:i.vendor,
    item_code:i.item_code,
    user_section:s.user_type
  }));

  try{
    await rest(
      'submitted_orders?on_conflict=order_date,product_name,user_section',
      'POST',
      lines,
      'resolution=merge-duplicates'
    );
  }catch(e){}

  const r=await fetch(
    SUPABASE_URL+'/functions/v1/send-order-notification',
    {
      method:'POST',
      headers:{
        'Content-Type':'application/json',
        'Authorization':'Bearer '+SUPABASE_ANON_KEY
      },
      body:JSON.stringify({
        submissionId:s.client_submission_id,
        submittedBy:s.user_type,
        orderDate:s.order_date,
        companyLogo:'https://raw.githubusercontent.com/pgonzalez2910/jengchi-product-images/main/jengchilogo.jpg',
        items:s.items
      })
    }
  );

  if(!r.ok)throw new Error(await r.text());

  const x=await r.json();

  if(!x.success){
    throw new Error(x.error||'Email failed');
  }

  s.status='synced';
  s.email_status='sent';
  s.synced_at=new Date().toISOString();

  await put('submissions',s);
}

async function syncPendingCatalogOrders(){
  const meta=await getAll('meta');

  const jobs=meta.filter(
    x=>
      String(x.key||'').startsWith('catalog-order-pending:') &&
      Array.isArray(x.value) &&
      x.value.length
  );

  for(const job of jobs){
    const failed=[];

    for(const row of job.value){
      try{
        await rest(
          `place_order_catalog?id=eq.${encodeURIComponent(row.id)}`,
          'PATCH',
          {
            sort_order:Number(row.sort_order),
            updated_at:new Date().toISOString()
          }
        );
      }catch(e){
        failed.push(row);
      }
    }

    await put('meta',{key:job.key,value:failed});
  }
}

async function syncPendingMasterSequences(){
  const meta=await getAll('meta');

  const jobs=meta.filter(
    x=>
      String(x.key||'').startsWith('master-sequence-pending:') &&
      Array.isArray(x.value) &&
      x.value.length
  );

  for(const job of jobs){
    const userType=String(job.key).split(':').pop();

    if(userType!=='rony'&&userType!=='julio')continue;

    try{
      await rest(
        'place_order_sequence_master?on_conflict=user_type',
        'POST',
        [{
          user_type:userType,
          sequence:job.value,
          updated_at:new Date().toISOString(),
          updated_by:'service-worker-build26'
        }],
        'resolution=merge-duplicates'
      );

      await put('meta',{key:job.key,value:[]});
    }catch(e){}
  }
}

async function backgroundSync(){
  await syncPendingMasterSequences();

  const drafts=(await getAll('drafts')).filter(x=>x.dirty);

  for(const d of drafts){
    try{
      await syncDraft(d);
    }catch(e){}
  }

  const subs=(await getAll('submissions')).filter(
    x=>x.status!=='synced'||x.email_status!=='sent'
  );

  for(const s of subs){
    try{
      await sendSubmission(s);
    }catch(e){}
  }

  const clients=await self.clients.matchAll({type:'window'});

  clients.forEach(c=>{
    c.postMessage({type:'SYNC_NOW'});
  });
}

self.addEventListener('sync',event=>{
  if(event.tag==='jc-place-order-sync'){
    event.waitUntil(backgroundSync());
  }
});
