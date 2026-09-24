// firebase-messaging-sw.js
//
// MUST be uploaded to your site's ROOT (same folder reception.html is served
// from, e.g. alongside manifest.json / icon-192.png) — Firebase requires the
// service worker's scope to cover the page that registers it, and a service
// worker's default scope is the folder it's served from.
//
// The Firebase project config is NOT hardcoded here (this file is static,
// can't read reception.html's localStorage) — reception.html passes it as
// query params when it calls navigator.serviceWorker.register(), e.g.
//   register('./firebase-messaging-sw.js?apiKey=...&projectId=...')
// This is safe: Firebase's client-side "apiKey" etc. are not secrets, they're
// meant to be public (this is standard practice, same as any Firebase web app).

importScripts('https://www.gstatic.com/firebasejs/10.13.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.13.1/firebase-messaging-compat.js');

var params = new URLSearchParams(self.location.search);
var firebaseConfig = {
  apiKey           : params.get('apiKey'),
  authDomain       : params.get('authDomain'),
  projectId        : params.get('projectId'),
  storageBucket    : params.get('storageBucket'),
  messagingSenderId: params.get('messagingSenderId'),
  appId            : params.get('appId'),
};

firebase.initializeApp(firebaseConfig);
var messaging = firebase.messaging();

// Fires when a push arrives while no tab has focus (or the browser is
// closed but the OS/browser keeps the service worker alive) — this is what
// actually shows a lock-screen / system-tray notification. When a tab IS
// focused, reception.html's own onMessage() handler in the page fires
// instead (see dlFcmEnable in the main HTML) — Firebase never delivers a
// message to both at once, so there's no double notification.
messaging.onBackgroundMessage(function (payload) {
  var n = payload.notification || {};
  var title = n.title || 'Payment Received';
  var body  = n.body  || '';
  self.registration.showNotification(title, {
    body : body,
    icon : './icon-192.png',
    badge: './icon-192.png',
    data : payload.data || {},
  });
});

// Clicking the notification focuses an already-open reception tab if there
// is one, otherwise opens a new one, instead of just closing silently.
self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
      for (var i = 0; i < list.length; i++) {
        if ('focus' in list[i]) return list[i].focus();
      }
      if (clients.openWindow) return clients.openWindow('./');
    })
  );
});
