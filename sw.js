// Cache para que el mapa siga funcionando sin señal despues de la primera
// carga del dia (con wifi/datos, por ejemplo en la mañana antes de salir).
//
// IMPORTANTE: cuando cambies app.js o style.css, sube este numero de
// version (ademas del ?v= en index.html) para que el celular note el
// cambio en vez de seguir usando la copia guardada.
var CACHE_VERSION = "mrd-v4";

var APP_SHELL = [
  "./",
  "./index.html",
  "./style.css?v=8",
  "./app.js?v=8",
  "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css",
  "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon.png",
  "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon-2x.png",
  "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(function (cache) {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (claves) {
      return Promise.all(
        claves
          .filter(function (clave) { return clave !== CACHE_VERSION; })
          .map(function (clave) { return caches.delete(clave); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (event) {
  var url = event.request.url;

  // Tiles del mapa (OpenStreetMap): una vez vistas, quedan guardadas para
  // siempre offline (no cambian, no hace falta revisar si hay una version
  // mas nueva).
  if (url.indexOf("tile.openstreetmap.org") !== -1) {
    event.respondWith(
      caches.open(CACHE_VERSION).then(function (cache) {
        return cache.match(event.request).then(function (guardado) {
          if (guardado) return guardado;
          return fetch(event.request).then(function (respuesta) {
            // Solo guarda tiles que si cargaron bien. Guardar un error de
            // paso (por ejemplo un corte breve de señal) lo dejaria
            // repitiendo ese error para siempre en esa zona del mapa.
            if (respuesta.ok) cache.put(event.request, respuesta.clone());
            return respuesta;
          });
        });
      })
    );
    return;
  }

  // Datos del dia (GitHub): primero intenta traer la version mas nueva; si
  // no hay señal, usa la ultima que se guardo con exito.
  if (url.indexOf("api.github.com") !== -1) {
    event.respondWith(
      fetch(event.request)
        .then(function (respuesta) {
          // IMPORTANTE: solo se guarda si la respuesta vino bien (200).
          // Si se guardara un error (por ejemplo un corte de señal a
          // mitad de una respuesta, o un 401/404 pasajero) y despues el
          // celular se queda sin señal, ese error quedaria repitiendose
          // como si fuera la respuesta real — pudiendo verse como "codigo
          // invalido" y borrando el acceso guardado sin motivo real.
          if (respuesta.ok) {
            var copia = respuesta.clone();
            caches.open(CACHE_VERSION).then(function (cache) {
              cache.put(event.request, copia);
            });
          }
          return respuesta;
        })
        .catch(function () {
          return caches.match(event.request);
        })
    );
    return;
  }

  // El resto (index.html, app.js, style.css, Leaflet): usa lo guardado y
  // de paso intenta traer una version mas nueva para la proxima vez.
  event.respondWith(
    caches.match(event.request).then(function (guardado) {
      var actualizar = fetch(event.request)
        .then(function (respuesta) {
          if (respuesta.ok) {
            caches.open(CACHE_VERSION).then(function (cache) {
              cache.put(event.request, respuesta.clone());
            });
          }
          return respuesta;
        })
        .catch(function () { return guardado; });
      return guardado || actualizar;
    })
  );
});
