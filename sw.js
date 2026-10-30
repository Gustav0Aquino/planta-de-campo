/* Guarda a casca para o app abrir sem sinal. Rede primeiro: com sinal, uma versão nova
   publicada já vale na próxima abertura; sem sinal, usa a última guardada.
   A planta de cada obra é guardada pelo app.js (cache "plantas"), não aqui. */
var CASCA = "casca-0.1.6";
var ARQUIVOS = ["./", "index.html", "app.js", "manifest.webmanifest", "icone-192.png", "icone-512.png", "vendor/msal-browser-4.28.1.min.js"];
self.addEventListener("install", function(e){
  e.waitUntil(caches.open(CASCA).then(function(c){ return c.addAll(ARQUIVOS); }).then(function(){ return self.skipWaiting(); }));
});
self.addEventListener("activate", function(e){
  e.waitUntil(caches.keys().then(function(ks){
    return Promise.all(ks.filter(function(k){ return k.indexOf("casca-") === 0 && k !== CASCA; }).map(function(k){ return caches.delete(k); }));
  }).then(function(){ return self.clients.claim(); }));
});
self.addEventListener("fetch", function(e){
  var u = new URL(e.request.url);
  if(e.request.method !== "GET" || u.origin !== location.origin) return;
  /* no-cache: confere com o servidor a cada abertura (sem isso o cache HTTP do navegador segurava a versão velha) */
  e.respondWith(fetch(e.request.url, { cache: "no-cache", credentials: "same-origin" }).then(function(r){
    if(r.ok){ var c = r.clone(); caches.open(CASCA).then(function(x){ x.put(e.request, c); }); }
    return r;
  }).catch(function(){
    return caches.match(e.request, { ignoreSearch: true }).then(function(r){ return r || caches.match("index.html"); });
  }));
});
