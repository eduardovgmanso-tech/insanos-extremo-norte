// Painel Insanos: busca sempre a versão nova; sem internet, abre a última versão vista.
const CACHE = "insanos-v9";
self.addEventListener("install", e => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;
  if (/\.mp4$/i.test(new URL(r.url).pathname)) return;      // vídeo de fundo: direto da rede (guardar vídeo no app trava a reprodução)
  e.respondWith(fetch(r, {cache: "no-store"}).then(resp => {
    if (resp.ok) { const c = resp.clone(); caches.open(CACHE).then(k => k.put(r, c)); }
    return resp;
  }).catch(() => caches.match(r, {ignoreSearch: true})));
});
self.addEventListener("push", e => {
  let d = {};
  try { d = e.data.json(); } catch (_) { d = {title: "Insanos ADM", body: e.data ? e.data.text() : ""}; }
  const op = {body: d.body || "", icon: d.icon || "icon-192.png", badge: "icon-192.png", tag: d.tag || "aviso", renotify: true,
              data: {url: d.url || "adm.html", zap: d.zap || ""}};
  if (d.zap) op.actions = [{action: "zap", title: "📲 Enviar no WhatsApp"}];
  // vibração própria de cada tipo de aviso (Android; o iPhone ignora): dá para reconhecer sem olhar o celular
  const VIB = {pag: [120, 90, 120], evento: [450], social: [450], pub: [450], insight: [150],
               cadastro: [200, 100, 200, 100, 200], niver: [100, 70, 100, 70, 100], aviso: [450, 120, 120]};
  const tipo = ((d.icon || "").match(/notif-([a-z]+)\.png/) || [])[1] || (d.tag === "pagamento" ? "pag" : d.tag === "aniversario" ? "niver" : d.tag === "afastamento" ? "aviso" : "");
  op.vibrate = VIB[tipo] || [200];
  e.waitUntil(self.registration.showNotification(d.title || "Insanos ADM", op));
});
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const dd = e.notification.data || {};
  // botão "Enviar no WhatsApp" (Android/computador) ou toque na notificação no iPhone (que não mostra botões):
  // abre o painel com a caixa da mensagem (o site do WhatsApp quebra os emojis; de lá o painel abre o app direto)
  if (dd.zap && (e.action === "zap" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {
    e.waitUntil(caches.open("insanos-zap").then(c => c.put("zap", new Response(JSON.stringify({ts: Date.now(), txt: dd.zap}))))
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
