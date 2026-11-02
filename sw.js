/* Golf Log service worker
   - Online: always fetches the newest files, so updates show up on the next launch.
   - Offline (no signal on the course): falls back to the saved copy, so the app still opens.
   Change CACHE to a new name on each release to be safe. */
const CACHE='golflog-2026-11-02-a';
const ASSETS=['./','./index.html','./manifest.json','./icon-180.png'];

self.addEventListener('install',e=>{
  e.waitUntil(
    caches.open(CACHE)
      .then(c=>Promise.all(ASSETS.map(a=>c.add(a).catch(()=>{}))))
      .then(()=>self.skipWaiting())
  );
});

self.addEventListener('activate',e=>{
  e.waitUntil(
    caches.keys()
      .then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener('fetch',e=>{
  const req=e.request;
  if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin!==location.origin)return;   // Sheets, fonts and API calls go straight to the network
  e.respondWith((async()=>{
    const cache=await caches.open(CACHE);
    try{
      const ctl=new AbortController();
      const t=setTimeout(()=>ctl.abort(),3000);   // weak signal: give up after 3s and use the saved copy
      const res=await fetch(req,{signal:ctl.signal,cache:'no-cache'});
      clearTimeout(t);
      if(res&&res.ok)cache.put(req,res.clone());
      return res;
    }catch(err){
      const hit=(await cache.match(req))||(await cache.match('./index.html'));
      if(hit)return hit;
      throw err;
    }
  })());
});
