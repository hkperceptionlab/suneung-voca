/* 수능 단어장 서비스 워커
   - 화면(index.html 등): 캐시를 먼저 보여 주고 뒤에서 새 버전을 받아 둠 → 다음에 열 때 반영
   - data/*.json: 네트워크 우선, 실패하면 캐시 → 새 Day가 올라오면 바로 반영
   - 글꼴: 한 번 받으면 캐시 사용
   화면 코드를 고치면 VERSION을 올려 주세요. (데이터만 바꿀 때는 그대로 둬도 됩니다) */
const VERSION = "v1";
const SHELL = "shell-" + VERSION, DATA = "data", FONTS = "fonts";
const SHELL_FILES = ["./", "index.html", "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil((async () => {
    await (await caches.open(SHELL)).addAll(SHELL_FILES);
    // 처음 설치할 때 지금 있는 Day 파일도 받아 둠 → 첫 방문 직후부터 오프라인 가능
    try {
      const data = await caches.open(DATA);
      const res = await fetch("data/days.json", { cache: "no-cache" });
      if (res.ok) {
        const days = await res.clone().json();
        await data.put("data/days.json", res);
        await data.addAll(days.flatMap(d => [d.words, d.csat].filter(Boolean).map(f => "data/" + f)));
      }
    } catch (err) {}
    self.skipWaiting();
  })());
});

self.addEventListener("activate", e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k.startsWith("shell-") && k !== SHELL) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin === location.origin && url.pathname.includes("/data/")) e.respondWith(networkFirst(req));
  else if (url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com") e.respondWith(cacheFirst(req, FONTS));
  else if (url.origin === location.origin) e.respondWith(staleWhileRevalidate(req));
});

async function networkFirst(req) {
  const cache = await caches.open(DATA);
  try {
    // no-cache: 브라우저 HTTP 캐시(GitHub Pages는 10분)를 건너뛰고 서버에 새 버전이 있는지 확인
    const res = await fetch(req, { cache: "no-cache" });
    if (res.ok) cache.put(req, res.clone());
    return res;
  } catch (err) {
    const hit = await cache.match(req, { ignoreSearch: true });
    if (hit) return hit;
    throw err;
  }
}

async function cacheFirst(req, name) {
  const cache = await caches.open(name);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok || res.type === "opaque") cache.put(req, res.clone());
  return res;
}

async function staleWhileRevalidate(req) {
  const cache = await caches.open(SHELL);
  const key = req.mode === "navigate" ? "index.html" : req;
  const hit = await cache.match(key, { ignoreSearch: req.mode === "navigate" });
  const fresh = fetch(req, { cache: "no-cache" }).then(res => {
    if (res.ok) cache.put(key, res.clone());
    return res;
  });
  if (hit) { fresh.catch(() => {}); return hit; }
  return fresh;
}
