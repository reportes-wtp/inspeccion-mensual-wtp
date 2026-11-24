const C='insp-v35',F=['./','index.html','manifest.json','icon-192.png','icon-512.png','https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(F)));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C&&x!=='insp-cfg').map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET'||e.request.url.includes('script.google'))return;
  const html=e.request.mode==='navigate'||e.request.url.endsWith('.html')||e.request.url.endsWith('/');
  e.respondWith(fetch(e.request,html?{cache:'no-store'}:{}).then(r=>{const cp=r.clone();caches.open(C).then(c=>c.put(e.request,cp));return r}).catch(()=>caches.match(e.request,{ignoreSearch:true})));
});
// ---------- subida de fotos en segundo plano (sigue aunque la pantalla se bloquee) ----------
function idb(){return new Promise((res,rej)=>{const q=indexedDB.open('insp_fotos',1);q.onupgradeneeded=()=>q.result.createObjectStore('f',{keyPath:'k'});q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)})}
function op(mode,fn){return idb().then(db=>new Promise((res,rej)=>{const t=db.transaction('f',mode),r=fn(t.objectStore('f'));t.oncomplete=()=>{res(r&&r.result);db.close()};t.onerror=()=>{rej(t.error);db.close()}}))}
const todas=()=>op('readonly',s=>s.getAll()).then(x=>x||[]);
const borrar=k=>op('readwrite',s=>s.delete(k));
async function b64(blob){const u=new Uint8Array(await blob.arrayBuffer());let s='';for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));return btoa(s)}
async function getUrl(){const r=await (await caches.open('insp-cfg')).match('/__url');return r?r.text():''}
async function setUrl(u){if(u)await (await caches.open('insp-cfg')).put('/__url',new Response(u))}
async function avisar(m){(await self.clients.matchAll({type:'window',includeUncontrolled:true})).forEach(c=>c.postMessage(m))}
let corriendo=null;
function subir(){if(!corriendo){const fn=()=>subirCola();corriendo=(self.navigator&&navigator.locks?navigator.locks.request('insp-fotos',fn):fn()).finally(()=>{corriendo=null})}return corriendo}
async function subirCola(){
  const url=await getUrl();if(!url)return{quedan:-1};
  const lista=(await todas()).sort((a,b)=>a.k<b.k?-1:1),total=lista.length;
  if(!total){await avisar({tipo:'fotos',hechas:0,total:0,quedan:0,fin:true});return{quedan:0}}
  const lotes=[];lista.forEach(f=>{const l=lotes[lotes.length-1];if(l&&l.id===f.id&&l.items.length<5)l.items.push(f);else lotes.push({id:f.id,carpeta:f.carpeta,items:[f]})});
  let idx=0,hechas=0,error=false;await avisar({tipo:'fotos',hechas,total,fin:false});
  async function w(){
    while(!error&&idx<lotes.length){const L=lotes[idx++];
      try{const fotos=[];for(const f of L.items)fotos.push({nombre:f.nombre,b64:await b64(f.blob)});
        const r=await fetch(url,{method:'POST',body:JSON.stringify({accion:'fotos',id:L.id,carpeta:L.carpeta,primero:L.items.some(f=>f.n<=5),fotos})});
        const j=await r.json();
        if(j&&j.ok){for(const f of L.items)await borrar(f.k);hechas+=L.items.length;await avisar({tipo:'fotos',hechas,total,fin:false})}else error=true;
      }catch(e){error=true}}
  }
  await Promise.all([w(),w(),w(),w()]);
  const quedan=(await todas()).length;await avisar({tipo:'fotos',hechas,total,quedan,fin:true});return{quedan};
}
self.addEventListener('message',e=>{const d=e.data||{};if(d.tipo==='subir')e.waitUntil(setUrl(d.url).then(subir))});
self.addEventListener('sync',e=>{if(e.tag==='insp-fotos')e.waitUntil(subir().then(r=>{if(r&&r.quedan>0)throw new Error('pendientes')}))});
