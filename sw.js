// Painel Insanos: busca sempre a versão nova; sem internet, abre a última versão vista.
const CACHE = "insanos-v6";
self.addEventListener("install", e => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;
  e.respondWith(fetch(r, {cache: "no-store"}).then(resp => {
    if (resp.ok) { const c = resp.clone(); caches.open(CACHE).then(k => k.put(r, c)); }
    return resp;
  }).catch(() => caches.match(r, {ignoreSearch: true})));
});
self.addEventListener("push", e => {
  let d = {};
  try { d = e.data.json(); } catch (_) { d = {title: "Insanos ADM", body: e.data ? e.data.text() : ""}; }
  const op = {body: d.body || "", icon: "icon-192.png", badge: "icon-192.png", tag: d.tag || "aviso", renotify: true,
              data: {url: d.url || "adm.html", zap: d.zap || ""}};
  if (d.zap) op.actions = [{action: "zap", title: "📲 Enviar no WhatsApp"}];
  e.waitUntil(self.registration.showNotification(d.title || "Insanos ADM", op));
});
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const dd = e.notification.data || {};
  // botão "Enviar no WhatsApp" (Android/computador) ou toque na notificação no iPhone (que não mostra botões):
  // abre o painel com a caixa da mensagem (o site do WhatsApp quebra os emojis; de lá o painel abre o app direto)
  if (dd.zap && (e.action === "zap" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {
    e.waitUntil(caches.open("insanos-zap").then(c => c.put("zap", new Response(JSON.stringify({ts: Date.now(), txt: dd.zap})))))
      .then(() => clients.matchAll({type: "window", includeUncontrolled: true}))
      .then(ws => { for (const w of ws) if (w.url.includes("adm.html")) { w.postMessage({zap: 1}); return w.focus(); }
                    return clients.openWindow("adm.html#zap"); }));
    return;
  }
  e.waitUntil(clients.matchAll({type: "window", includeUncontrolled: true}).then(ws => {
    for (const w of ws) if (w.url.includes("adm.html")) return w.focus();
    return clients.openWindow((e.notification.data && e.notification.data.url) || "adm.html");
  }));
});
