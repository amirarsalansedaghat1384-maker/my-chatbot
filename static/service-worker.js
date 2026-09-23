// سرویس‌ورکر ساده فقط برای فعال شدن قابلیت نصب (Add to Home Screen)
const CACHE_NAME = "chatbot-cache-v1";

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  self.clients.claim();
});

// یه پاسخ‌دهنده‌ی ساده fetch (بدون کش پیچیده) - فقط برای معتبر بودن PWA لازمه
self.addEventListener("fetch", (event) => {
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
