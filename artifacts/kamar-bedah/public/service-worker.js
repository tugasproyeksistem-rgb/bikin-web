// Service Worker — Sistem Kamar Bedah RS Panti Rini
// Tujuan: cache SHELL aplikasi (HTML/JS/CSS/logo) agar load lebih cepat
// dan tetap bisa dibuka saat sinyal lemah/putus sebentar.
// TIDAK menyimpan/cache data Supabase — data selalu live dari server.

const CACHE_NAME = "kamar-bedah-shell-v1";
const SHELL_ASSETS = [
  "/",
  "/index.html",
  "/manifest.json",
  "/logo.jpeg",
  "/favicon.svg",
];

// Saat install: simpan shell dasar ke cache
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS))
  );
  self.skipWaiting();
});

// Saat activate: bersihkan cache versi lama (kalau ada update CACHE_NAME)
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// Strategi fetch: NETWORK-FIRST untuk HTML/JS/CSS (selalu coba ambil versi
// terbaru dulu saat online — supaya update kode/deploy baru langsung terlihat),
// baru fallback ke cache kalau request gagal (sinyal lemah/offline).
// Request ke Supabase (*.supabase.co) TIDAK pernah di-intercept — biarkan
// langsung ke network seperti biasa, supaya data selalu realtime.
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Jangan campur tangani request API/Supabase/Dropbox — biarkan lewat langsung
  if (
    url.hostname.includes("supabase.co") ||
    url.hostname.includes("dropboxapi.com") ||
    event.request.method !== "GET"
  ) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        // Simpan salinan terbaru ke cache untuk fallback offline nanti
        const resClone = networkResponse.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
        return networkResponse;
      })
      .catch(() => {
        // Network gagal (offline/sinyal putus) → coba ambil dari cache
        return caches.match(event.request).then((cached) => {
          return cached || caches.match("/index.html");
        });
      })
  );
});
