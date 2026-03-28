// 가계부 Service Worker v1.0
// 역할: 오프라인 동작 보장 + 앱 파일 캐시

const CACHE_NAME = 'gaegyebu-v2';
const ASSETS = [
  './index.html',
  './budget-data.js',
  './manifest.json',
  'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js'
];

// 설치: 핵심 파일 캐시
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      // SheetJS는 실패해도 무시 (CDN 불가 시)
      return cache.addAll([
        './index.html',
        './budget-data.js',
        './manifest.json'
      ]);
    })
  );
  self.skipWaiting();
});

// 활성화: 구버전 캐시 정리
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// 요청 가로채기: 캐시 우선, 없으면 네트워크
self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        // 앱 파일만 캐시 갱신
        if (e.request.url.includes('Geminai') || e.request.url.includes('budget-data')) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => cached); // 네트워크 실패 시 캐시 반환
    })
  );
});
