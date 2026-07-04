// ══════════════════════════════════════════════════════════
// NABS&FBK HOME OS — Service Worker v2.0
// Cache Strategy: Cache-First for static assets
//                 Network-First for API calls
// ══════════════════════════════════════════════════════════

const CACHE_NAME    = 'nabsfbk-home-os-v2';
const CACHE_DYNAMIC = 'nabsfbk-dynamic-v2';

// Static assets to cache immediately on install
const STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Hind+Siliguri:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap'
];

// ─── INSTALL ─────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('[SW] Pre-caching static assets');
      return cache.addAll(STATIC_ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// ─── ACTIVATE ────────────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys
          .filter(k => k !== CACHE_NAME && k !== CACHE_DYNAMIC)
          .map(k => {
            console.log('[SW] Deleting old cache:', k);
            return caches.delete(k);
          })
      );
    }).then(() => self.clients.claim())
  );
});

// ─── FETCH ───────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Skip non-GET and chrome-extension requests
  if (event.request.method !== 'GET') return;
  if (url.protocol === 'chrome-extension:') return;

  // API calls (Apps Script, RSS feeds, YouTube) → Network-First
  if (
    url.hostname.includes('script.google.com') ||
    url.hostname.includes('googleapis.com') ||
    url.hostname.includes('rss2json.com') ||
    url.hostname.includes('feeds.bbci.co.uk') ||
    url.hostname.includes('feeds.reuters.com') ||
    url.hostname.includes('prothomalo.com') ||
    url.hostname.includes('ip-api.com')
  ) {
    event.respondWith(networkFirst(event.request));
    return;
  }

  // Google Fonts → Cache-First (they don't change often)
  if (url.hostname.includes('fonts.googleapis.com') || url.hostname.includes('fonts.gstatic.com')) {
    event.respondWith(cacheFirst(event.request));
    return;
  }

  // Everything else → Stale-While-Revalidate (fast + fresh)
  event.respondWith(staleWhileRevalidate(event.request));
});

// ─── STRATEGIES ──────────────────────────────────────────

async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response('Offline — cache miss', { status: 503 });
  }
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_DYNAMIC);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    return cached || new Response(JSON.stringify({ status: 'offline', msg: 'No network' }),
      { status: 503, headers: { 'Content-Type': 'application/json' } });
  }
}

async function staleWhileRevalidate(request) {
  const cache    = await caches.open(CACHE_NAME);
  const cached   = await cache.match(request);
  const fetchPromise = fetch(request).then(response => {
    if (response.ok) cache.put(request, response.clone());
    return response;
  }).catch(() => cached);
  return cached || fetchPromise;
}

// ─── BACKGROUND SYNC (for offline entry queue) ───────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-entries') {
    event.waitUntil(syncPendingEntries());
  }
});

async function syncPendingEntries() {
  // When back online, this fires automatically
  // Frontend queues entries in localStorage under 'nabsfbk_pending_sync'
  console.log('[SW] Background sync triggered for pending entries');
  // Notify all open clients
  const clients = await self.clients.matchAll({ type: 'window' });
  clients.forEach(client => client.postMessage({ type: 'SYNC_NOW' }));
}

// ─── PUSH NOTIFICATIONS (future use) ─────────────────────
self.addEventListener('push', event => {
  if (!event.data) return;
  const data = event.data.json();
  event.waitUntil(
    self.registration.showNotification(data.title || 'NABS&FBK Home OS', {
      body:  data.body  || '',
      icon:  './icon-192.png',
      badge: './icon-192.png',
      tag:   data.tag   || 'nabsfbk-notification',
      data:  data.url   || './'
    })
  );
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.openWindow(event.notification.data || './')
  );
});
