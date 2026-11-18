// sudappka – minimalny service worker (pozwala zainstalować aplikację na telefonie).
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => self.clients.claim());
self.addEventListener('fetch', () => {});
