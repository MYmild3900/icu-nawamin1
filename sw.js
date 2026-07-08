// Service Worker — ระบบพัสดุ ICU นวมินทร์ 1
// ทำให้แอปเปิดได้แม้ไม่มีเน็ต (cache ตัวแอป + ไลบรารีในเครื่อง)
// กลยุทธ์: network-first สำหรับไฟล์แอป (ออนไลน์ได้เวอร์ชันล่าสุด, ออฟไลน์ใช้ที่ cache ไว้)
// ส่วน Google Sheets / Apps Script (คนละ origin) ปล่อยไปเน็ตตามปกติ — ออฟไลน์แล้วระบบคิวจัดการเอง

const CACHE = 'icu-nawamin1-v1';
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './qrcode.min.js',
  './jsQR.min.js',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
      .catch(function () { /* ไฟล์บางตัวโหลดไม่ได้ตอนติดตั้ง — ข้ามไป */ })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.filter(function (k) { return k !== CACHE; })
          .map(function (k) { return caches.delete(k); }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url;
  try { url = new URL(req.url); } catch (err) { return; }

  // จัดการเฉพาะไฟล์ของแอป (same-origin) — ปล่อย Google Sheets/Apps Script/รูป CDN ไปเน็ตปกติ
  if (url.origin !== self.location.origin) return;

  e.respondWith(
    fetch(req)
      .then(function (res) {
        // ออนไลน์ได้ไฟล์ใหม่ → อัปเดต cache ไว้ใช้ตอนออฟไลน์
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); }).catch(function () {});
        return res;
      })
      .catch(function () {
        // ออฟไลน์ → ใช้ไฟล์จาก cache; ถ้าเป็นการเปิดหน้า ให้ตกไปที่ index.html
        return caches.match(req).then(function (r) {
          return r || caches.match('./index.html');
        });
      })
  );
});
