// Service worker del sitio de presentaciones: permite presentar sin conexión.
//
// - Las imágenes llevan ?v=<versión> en la dirección; como una versión nunca cambia de contenido,
//   se sirven desde la caché si están (y se guardan la primera vez que se ven).
// - Páginas, visor.json y estilos: se intenta la red primero (para ver siempre lo último) y, si no hay
//   conexión o tarda demasiado, se usa la copia guardada.
// - Los PDF y PowerPoint de descarga no se guardan (son pesados y no hacen falta para presentar).
// - El botón "Guardar para usar sin conexión" de cada ficha descarga todo de una vez a esta misma caché.

const CACHE = 'presentaciones-v1';
const ESPERA_RED_MS = 3500;   // con wifi lento, pasado este tiempo se usa la copia guardada

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

function guardable(url, respuesta) {
  return respuesta && respuesta.ok && respuesta.type === 'basic' && !/\.(pdf|pptx?|odp)$/i.test(url.pathname);
}

async function primeroCache(peticion, url) {
  const cache = await caches.open(CACHE);
  const guardada = await cache.match(peticion);
  if (guardada) return guardada;
  const respuesta = await fetch(peticion);
  if (guardable(url, respuesta)) cache.put(peticion, respuesta.clone());
  return respuesta;
}

async function primeroRed(peticion, url) {
  const cache = await caches.open(CACHE);
  const guardada = await cache.match(peticion, { ignoreSearch: url.pathname.endsWith('.css') });
  const red = fetch(peticion).then(respuesta => {
    if (guardable(url, respuesta)) cache.put(peticion, respuesta.clone());
    return respuesta;
  });
  if (!guardada) return peticion.mode === 'navigate' ? red.catch(() => paginaSinConexion()) : red;
  // Si hay copia, no esperar indefinidamente a una red lenta
  const plazo = new Promise(resolver => setTimeout(() => resolver(guardada), ESPERA_RED_MS));
  return Promise.race([red.catch(() => guardada), plazo]);
}

function paginaSinConexion() {
  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Sin conexión</title>
<style>body{font-family:system-ui,sans-serif;background:#0f1115;color:#e8eaee;display:flex;align-items:center;
justify-content:center;min-height:100vh;margin:0;padding:24px;box-sizing:border-box;text-align:center}
a{color:#4da3ff}div{max-width:460px}h1{font-size:1.4rem}p{color:#9aa3b0;line-height:1.5}</style></head>
<body><div><h1>Sin conexión</h1><p>Esta página no está guardada en este dispositivo.
Para presentar sin internet, abre la ficha de la presentación con conexión y pulsa
«Guardar para usar sin conexión».</p><p><a href="javascript:history.back()">← Volver</a></p></div></body></html>`;
  return new Response(html, { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } });
}

self.addEventListener('fetch', e => {
  const peticion = e.request;
  if (peticion.method !== 'GET') return;
  const url = new URL(peticion.url);
  if (url.origin !== self.location.origin) return;
  if (/\.(pdf|pptx?|odp)$/i.test(url.pathname)) return;   // descargas: directo a la red
  e.respondWith(url.searchParams.has('v') ? primeroCache(peticion, url) : primeroRed(peticion, url));
});
