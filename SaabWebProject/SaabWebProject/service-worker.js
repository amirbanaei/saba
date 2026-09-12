


const CACHE_NAME = "my-app-cache-v2";

/* ---------- فایل‌های استاتیک ---------- */
const STATIC_ASSETS = [
    "/",
    "/offline.html",
    "/Scripts/jquery-3.3.1.min.js",
    "/Scripts/Usertest.js",
    "/Content/site.css"
];

/* ---------- صفحات MVC که باید آفلاین باشند ---------- */
const OFFLINE_PAGES = [
    "/Setting/Orders/viewEntezar",
    "/Setting/Orders/Usertest",
    "/Setting/Orders/seeerhaj",
    "/Setting/Orders/seeerhaj2",
    "/Orders/erja_sabt_usr1"   // 👈 اضافه شد
];


/* ================= INSTALL ================= */
self.addEventListener("install", event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(async cache => {
            const allPages = [...STATIC_ASSETS, ...OFFLINE_PAGES];

            for (const url of allPages) {
                try {
                    await cache.add(url);
                } catch (e) {
                    console.warn("Cache failed:", url);
                }
            }
        })
    );
    self.skipWaiting();
});


/* ================= ACTIVATE ================= */
self.addEventListener("activate", event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys.map(k => (k !== CACHE_NAME ? caches.delete(k) : null))
            )
        )
    );
    self.clients.claim();
});


/* ================= FETCH ================= */
self.addEventListener("fetch", event => {
    const req = event.request;

    if (req.method !== "GET") return;

    /* ---------- JSON API ---------- */
    if (req.headers.get("Accept")?.includes("application/json")) {
        event.respondWith(
            fetch(req).then(res => {
                const clone = res.clone();
                caches.open(CACHE_NAME).then(c => c.put(req, clone));
                return res;
            }).catch(() => caches.match(req))
        );
        return;
    }

    /* ---------- PAGE NAVIGATION ---------- */
    if (req.mode === "navigate") {
        event.respondWith(
            fetch(req).then(res => {
                const clone = res.clone();
                caches.open(CACHE_NAME).then(c => c.put(req, clone));
                return res;
            }).catch(async () => {

                const cache = await caches.open(CACHE_NAME);

                const matchedPage = OFFLINE_PAGES.find(p =>
                    req.url.includes(p)
                );

                if (matchedPage) {
                    return cache.match(matchedPage, { ignoreSearch: true });
                }

                return cache.match("/offline.html");
            })
        );
        return;
    }

    /* ---------- STATIC FILES ---------- */
    event.respondWith(
        fetch(req).then(res => {
            if (!res || res.status !== 200 || res.type === "opaque") {
                return res;
            }
            const clone = res.clone();
            caches.open(CACHE_NAME).then(c => c.put(req, clone));
            return res;
        }).catch(() => caches.match(req))
    );
});
