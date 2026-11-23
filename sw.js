/* Service worker RS Hermina Balikpapan: cepat (cache), tetap jalan saat sinyal lemah, notifikasi. Naikkan V bila mengganti file. */
const V='hm-v1',SHELL=['./','icon-192.png','icon-512.png','manifest.webmanifest'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V&&x!=='hm-cfg').map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
function nf(r,ms,fb){return new Promise(res=>{let done=false;const ok=x=>{if(!done){done=true;res(x);}};
  const t=setTimeout(async()=>{const c=await caches.match(r);if(c)ok(c);},ms);
  fetch(r).then(async n=>{clearTimeout(t);if(n.ok){const c=await caches.open(V);c.put(r,n.clone());}ok(n);})
   .catch(async()=>{clearTimeout(t);ok(await caches.match(r)||(fb&&await caches.match(fb))||Response.error());});});}
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;const u=new URL(r.url);
  if(u.searchParams.has('action'))return e.respondWith(nf(r,7000));
  if(r.mode==='navigate')return e.respondWith(nf(r,4000,'./'));
  if(u.origin===location.origin||/unpkg\.com|cdnjs\.cloudflare\.com|fonts\./.test(u.hostname))
    e.respondWith(caches.match(r).then(c=>c||fetch(r).then(n=>{if(n.ok||n.type==='opaque'){const k=n.clone();caches.open(V).then(x=>x.put(r,k));}return n;})));});
self.addEventListener('message',e=>{const d=e.data||{};if(d.t==='cfg')caches.open('hm-cfg').then(c=>c.put('api',new Response(d.api)));if(d.t==='seen')caches.open('hm-cfg').then(c=>c.put('last',new Response(d.id)));});
self.addEventListener('periodicsync',e=>{if(e.tag==='hm-notif')e.waitUntil(chk());});
async function chk(){const c=await caches.open('hm-cfg'),a=await c.match('api');if(!a)return;const api=await a.text(),l=await c.match('last'),last=l?await l.text():'';
  const j=await (await fetch(api+'?action=getNotifs')).json();if(!j.success||!j.data.length)return;const n=j.data[0];
  if(n.id>last){await c.put('last',new Response(n.id));if(last)await self.registration.showNotification(n.title,{body:n.body||'Ada pembaruan di katalog RS Hermina Balikpapan',icon:'icon-192.png',badge:'icon-192.png',tag:n.id,data:{url:'./'}});}}
self.addEventListener('notificationclick',e=>{e.notification.close();e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(l=>{for(const w of l)if('focus' in w)return w.focus();return self.clients.openWindow('./');}));});
