// Painel Insanos: busca sempre a versão nova; sem internet, abre a última versão vista.
const CACHE = "insanos-v13";
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
  if (d.link) { op.data.link = d.link; op.actions = [{action: "link", title: d.lrot || "Abrir"}]; }
  if (d.rsvp) { op.data.rsvp = d.rsvp; op.actions = [{action: "vou", title: "✅ Vou"}, {action: "nao", title: "❌ Não vou"}]; op.requireInteraction = true; }
  op.data.para = d.para || ""; op.data.ru = d.ru || ""; op.data.tt = d.title || "";
  if (d.sos) {   // SOS: distância aproximada pela última localização guardada neste celular, e botões
    op.data.sosUrl = d.url || "adm.html"; op.requireInteraction = true; op.vibrate = [700, 200, 700, 200, 700];
    const S = d.sos, temLoc = S.lat !== "" && S.lat != null;
    if (temLoc) op.data.mapa = "https://www.google.com/maps/search/?api=1&query=" + S.lat + "," + S.lng;
    op.actions = [{action: "indo", title: "🏍️ Estou indo"}].concat(temLoc ? [{action: "mapa", title: "📍 Ver no mapa"}] : []);
    const comDist = temLoc ? idbGet("loc").then(l => {
      if (!l) return;
      const R = 6371, x = (S.lat - l.lat) * Math.PI / 180, y = (S.lng - l.lng) * Math.PI / 180;
      const h = Math.sin(x / 2) ** 2 + Math.cos(l.lat * Math.PI / 180) * Math.cos(S.lat * Math.PI / 180) * Math.sin(y / 2) ** 2, km = 2 * R * Math.asin(Math.sqrt(h));
      const velha = Date.now() - l.ts > 12 * 3600e3;
      op.body += " · a ~" + (km < 1 ? Math.round(km * 1000) + " m" : km.toFixed(1).replace(".", ",") + " km") + " de você" + (velha ? " (localização antiga)" : "");
    }).catch(() => {}) : Promise.resolve();
    e.waitUntil(Promise.all([comDist.then(() => self.registration.showNotification(d.title || "🆘 SOS", op)), registrarRecebeu(d)]));
    return;
  }
  // vibração própria de cada tipo de aviso (Android; o iPhone ignora): dá para reconhecer sem olhar o celular
  const VIB = {pag: [120, 90, 120], evento: [450], social: [450], pub: [450], insight: [150],
               cadastro: [200, 100, 200, 100, 200], niver: [100, 70, 100, 70, 100], aviso: [450, 120, 120]};
  const tipo = ((d.icon || "").match(/notif-([a-z]+)\.png/) || [])[1] || (d.tag === "pagamento" ? "pag" : d.tag === "aniversario" ? "niver" : d.tag === "afastamento" ? "aviso" : "");
  op.vibrate = VIB[tipo] || [200];
  // controle de entrega: registra na aba Acessos da planilha que o aviso chegou a este celular
  e.waitUntil(Promise.all([self.registration.showNotification(d.title || "Insanos ADM", op), registrarRecebeu(d)]));
});
function registrarRecebeu(d) {
  return (d.ru && d.para) ? fetch(d.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: d.para, ok: true,
      painel: "🔔 Recebeu: " + (d.title || ""), aparelho: aparelhoSW()})}).catch(() => {}) : Promise.resolve();
}
// última localização guardada neste celular (pela tela de entrada ou pelo Meu painel), para a distância do SOS
function idbGet(k) {
  return new Promise(r => { try { const o = indexedDB.open("insanos", 1); o.onupgradeneeded = () => o.result.createObjectStore("kv");
    o.onsuccess = () => { try { const q = o.result.transaction("kv").objectStore("kv").get(k); q.onsuccess = () => r(q.result || null); q.onerror = () => r(null); } catch (e) { r(null); } };
    o.onerror = () => r(null); } catch (e) { r(null); } });
}
function aparelhoSW() { const u = navigator.userAgent || ""; return /iPhone|iPad|iPod/.test(u) ? "iPhone" : /Android/.test(u) ? "Celular Android" : "Computador"; }
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const dd = e.notification.data || {};
  if (dd.ru && dd.para)    // controle: o irmão tocou no aviso
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
      painel: "👆 Abriu: " + (dd.tt || "") + (e.action === "link" ? " (" + (e.notification.actions && e.notification.actions[0] ? e.notification.actions[0].title : "botão") + ")" : ""),
      aparelho: aparelhoSW()})}).catch(() => {}));
  // botão "Enviar no WhatsApp" (Android/computador) ou toque na notificação no iPhone (que não mostra botões):
  // abre o painel com a caixa da mensagem (o site do WhatsApp quebra os emojis; de lá o painel abre o app direto)
  if (dd.sosUrl) {   // SOS: "Ver no mapa" abre o mapa; "Estou indo" ou o toque abrem a tela com a distância exata
    if (e.action === "mapa" && dd.mapa) { e.waitUntil(clients.openWindow(dd.mapa)); return; }
    e.waitUntil(clients.openWindow(dd.sosUrl + (e.action === "indo" ? "&indo=1" : ""))); return;
  }
  if (dd.rsvp && (e.action === "vou" || e.action === "nao") && dd.ru && dd.para) {   // respondeu direto na notificação
    const r = dd.rsvp, vou = e.action === "vou";
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
        painel: "✋|" + e.action + "|" + r.d + "|" + r.tk + "|" + r.e, aparelho: aparelhoSW()})}).catch(() => {})
      .then(() => self.registration.showNotification(vou ? "✅ Presença confirmada" : "👍 Resposta registrada",
        {body: vou ? "Valeu, irmão! Você vai na " + r.e + ". 🤝" : "Tranquilo, irmão. Registramos que você não vai na " + r.e + ".",
         icon: "icon-192.png", badge: "icon-192.png", tag: "rsvp-ok"})));
    return;
  }
  if (dd.rsvp && !e.action) {   // tocou na notificação (no iPhone não há botões): abre o painel para responder
    e.waitUntil(clients.openWindow("adm.html#eventos")); return;
  }
  if (dd.link && (e.action === "link" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {   // ex.: "Abrir meu boleto"
    e.waitUntil(clients.openWindow(dd.link)); return;
  }
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
