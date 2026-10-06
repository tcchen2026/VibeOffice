/* VibeOffice service worker. Registered by common/suite.js from every page, scope public/.
 *
 * Dev mode: no fetch handler, so every request goes to the server as if there were no worker, and
 * an edit shows on the next reload. On activation it deletes every cache an earlier worker left.
 *
 * Offline support comes later, as a versioned precache (each app's index.html and the scripts it
 * lists, the start page, manifest and icons) with an offline fallback for navigations.
 */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((names) => Promise.all(names.map((n) => caches.delete(n)))).then(() => self.clients.claim()));
});
