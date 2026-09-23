# Mapa de Rutas Diarias

Reemplazo liviano de Google My Maps para publicar la ruta diaria de cada camión.
Usa OpenStreetMap + Leaflet en vez de Google Maps, así que consume muchos menos
datos móviles al abrirlo — importante para los choferes que no tienen plan de datos.

- **Sitio publicado**: https://coordinadorl.github.io/mapa-rutas-diarias/
- **Repo de este sitio** (público, sin datos de clientes): `mapa-rutas-diarias`
- **Repo de datos** (privado, con clientes/direcciones/ventas reales): `mapa-rutas-datos`

## Por qué dos repos

Este repo (`mapa-rutas-diarias`) es **público** para que GitHub Pages sea
gratis, así que no puede tener datos reales de clientes. Esos datos viven en
`mapa-rutas-datos`, que es **privado**. El mapa público los lee usando un
"código de acceso" (un token de ese repo privado) — ver [TOKEN.md](TOKEN.md)
para generarlo y repartirlo.

## Choferes vs. administrador

Hay dos niveles de acceso, cada uno con su propio código (ver TOKEN.md):

- **Chofer**: solo ve el mapa y los clientes de su camión. No puede editar nada.
- **Administrador** (vos): además de ver el mapa, podés **arrastrar un punto**
  para corregir la ubicación de un cliente. Se guarda solo (abajo aparece
  "Ubicacion guardada") y desde ese momento todos los choferes ven el punto
  ya corregido — sin que tengan que hacer nada de su lado.

Para no tener que pegar el código a mano en cada celular (es largo), se
reparte como un **link** que ya lo trae incluido — ver la sección
"Como armar los links" en [TOKEN.md](TOKEN.md).

Como administrador tenés un botón flotante **"Compartir con choferes"**
(solo vos lo ves) donde pegás una vez el código de chofer y después podés:
- Abrir WhatsApp con el mensaje y el link ya armados, listo para elegir a
  quién mandárselo.
- Generar un código QR para imprimir o mostrar en pantalla — un chofer
  nuevo lo escanea con la cámara y queda configurado, sin copiar nada.

## Colores por vendedor

Dentro de cada camión, los clientes se pintan de un color distinto según su
código de vendedor (siempre colores bien diferenciables entre sí — nunca dos
tonos parecidos en el mismo camión). Si un camión tiene más de un vendedor,
aparece una leyenda de colores debajo del resumen. Los colores se pueden
repetir de un camión a otro, pero nunca dentro del mismo.

## Navegación y etiquetas

Cada cliente tiene en su popup dos botones chiquitos, **Maps** y **Waze**,
que abren la navegación directa a ese punto (usan la app si está instalada
en el celular). Además, al acercar bastante el zoom aparece el nombre de
cada cliente al lado del punto (como en My Maps); al alejar, se ocultan
solas para no saturar el mapa.

## Para el chofer en la calle

- **Buscar cliente** (🔍 abajo a la izquierda): filtra por nombre dentro del
  camión seleccionado y al tocar un resultado salta directo a ese punto y
  abre su popup.
- **Mi ubicación** (📍 arriba a la izquierda, junto al zoom): muestra un
  punto azul con la posición del celular, para ver qué tan cerca está de
  las próximas paradas. Nunca se manda a ningún lado, solo se dibuja en
  ese celular.
- **Marcar entregado**: botón dentro del popup de cada cliente. El punto
  queda atenuado con una palomita, y arriba se cuenta cuántos van
  entregados. Se guarda solo en el celular de quien lo marca y se
  reinicia solo al otro día con la ruta nueva (no se sube a ningún repo).

## Modo sin conexión

El sitio guarda en el celular (service worker, `sw.js`) el mapa base y las
imágenes del mapa que ya se vieron, además de la última versión de los
datos del día. Así, después de abrirlo una vez con wifi o datos (por
ejemplo en la mañana antes de salir), sigue funcionando aunque después se
quede sin señal — no puede descargar zonas nuevas del mapa sin conexión,
pero todo lo ya visto y los datos del día quedan disponibles.

**Importante**: cuando cambies `app.js` o `style.css`, además de subir el
número de `?v=` en `index.html` (ver más abajo), sube también
`CACHE_VERSION` en `sw.js` — si no, el celular puede seguir usando la
copia vieja guardada.

## Cómo se usa cada día

1. Genera los archivos `MAPA F1xx.xlsx` de siempre en la carpeta del día
   (ej: `C:\RUTAS DIARIAS\15MARTES 15`), igual que hasta ahora para Google My Maps.
2. Arrastra esa carpeta sobre `actualizar_ruta.bat` (o ejecútalo y pega la ruta
   cuando la pida). Esto genera `data/latest.json` local con los datos del día.
3. Corre `publicar_datos.bat`. Copia ese archivo al repo privado
   `mapa-rutas-datos` y lo sube.
4. Los choferes abren siempre el mismo link (arriba) y ven la ruta
   actualizada, sin instalar ni configurar nada (salvo pegar el código de
   acceso la primera vez).

También se puede correr todo a mano:
```
python convertir.py "C:\RUTAS DIARIAS\15MARTES 15"
```

## Estructura del proyecto

- `index.html`, `app.js`, `style.css` — el visor del mapa (selector de camión,
  mapa, pantalla de código de acceso).
- `convertir.py` — convierte los `MAPA F1xx.xlsx` del día a `data/latest.json`.
- `data/latest.json` — datos del día actual, solo en esta computadora
  (está en `.gitignore`, nunca se sube a este repo público).
- `actualizar_ruta.bat` — genera `data/latest.json` desde la carpeta del día.
- `publicar_datos.bat` — sube ese archivo al repo privado `mapa-rutas-datos`.
- `TOKEN.md` — cómo generar y repartir el código de acceso.
- `sw.js` — service worker para que el mapa funcione sin conexión.

El repo privado `mapa-rutas-datos` es una carpeta aparte en
`C:\MAPA-RUTAS-DATOS` (clon de ese repo); `publicar_datos.bat` ya sabe copiar
ahí el archivo y subirlo. Ese mismo repo privado también guarda
`overrides.json`, que crea y actualiza el mapa solo cuando el administrador
corrige la ubicación de un cliente (no hace falta tocarlo a mano).

## Probar en local

```
python -m http.server 8532
```
y abrir `http://localhost:8532` en el navegador. Sin código de acceso válido
va a mostrar la pantalla para pedirlo (es normal, esa parte solo funciona con
un código real generado según TOKEN.md).

## Si cambias `app.js` o `style.css`

Los celulares los guardan en caché, así que un cambio puede tardar en
notarse. Sube el cambio subiendo también el número en
`index.html` (`app.js?v=2` → `?v=3`, etc.) para forzar que se note la
actualización.

## Publicar en GitHub Pages (ya hecho para este repo)

1. Crear el repositorio en GitHub y hacer push de esta carpeta.
2. En **Settings → Pages**, elegir la rama `main` y carpeta raíz (`/`).
3. GitHub entrega un link fijo (`https://usuario.github.io/repo/`) que es el
   que se comparte con los choferes.
