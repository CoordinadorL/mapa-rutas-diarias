# Mapa de Rutas Diarias

Reemplazo liviano de Google My Maps para publicar la ruta diaria de cada camión.
Usa OpenStreetMap + Leaflet en vez de Google Maps, así que consume muchos menos
datos móviles al abrirlo — importante para los choferes que no tienen plan de datos.

## ⚠️ Pendiente: cómo va a llegar `data/latest.json` a los choferes

Este repo es **público** (para que GitHub Pages sea gratis), y `data/latest.json`
trae nombres, direcciones y ventas de clientes reales. Por eso **no se sube a
git** (está en `.gitignore`) — si se sube, cualquier persona en internet podría
verlo.

Eso significa que, tal como está ahora, el sitio publicado en GitHub Pages
**no va a tener datos** (mostrará el error "no se pudo cargar la ruta del día"),
porque el archivo que los contiene nunca llega al repo público. Falta decidir
cómo se entrega ese archivo a diario sin exponerlo, por ejemplo:

- Un repo **privado** aparte solo para `data/latest.json`, y que `app.js` lo
  lea desde ahí (requiere que ese repo/host sí soporte acceso privado).
- Guardarlo en un Google Drive/OneDrive con link no listado y que `app.js`
  lo lea de ahí en vez de `data/latest.json`.
- Pasar `data/latest.json` a los choferes de otra forma (ej. adjunto en el
  grupo de WhatsApp del turno) en vez de servirlo desde la página.

Cuando se decida, hay que ajustar la URL del `fetch(...)` en `app.js`.

## Cómo se usa cada día (una vez resuelto lo anterior)

1. Genera los archivos `MAPA F1xx.xlsx` de siempre en la carpeta del día
   (ej: `C:\RUTAS DIARIAS\15MARTES 15`), igual que hasta ahora para Google My Maps.
2. Arrastra esa carpeta sobre `actualizar_ruta.bat` (o ejecútalo y pega la ruta
   cuando la pida). Esto genera `data/latest.json` con los datos del día
   (queda solo en tu computadora, no se sube a este repo).
3. Entrega ese archivo a donde corresponda según lo que se haya definido arriba.
4. Los choferes abren siempre el mismo link (la página publicada en GitHub
   Pages) y ven la ruta actualizada, sin tener que instalar ni configurar nada.

También se puede correr directo:
```
python convertir.py "C:\RUTAS DIARIAS\15MARTES 15"
```

## Estructura del proyecto

- `index.html`, `app.js`, `style.css` — el visor del mapa (selector de camión + mapa).
- `convertir.py` — convierte los `MAPA F1xx.xlsx` del día a `data/latest.json`.
- `data/latest.json` — datos del día actual (se sobrescribe cada vez que se corre el
  conversor). **No está en git** porque trae datos reales de clientes; ver la
  sección de arriba.
- `actualizar_ruta.bat` — atajo para no tener que escribir el comando de Python.

## Probar en local antes de publicar

```
python -m http.server 8532
```
y abrir `http://localhost:8532` en el navegador.

## Publicar en GitHub Pages

1. Crear el repositorio en GitHub y hacer push de esta carpeta.
2. En **Settings → Pages**, elegir la rama `main` y carpeta raíz (`/`).
3. GitHub entrega un link fijo (tipo `https://usuario.github.io/mapa-rutas/`)
   que es el que se comparte con los choferes.
