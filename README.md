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
  las próximas paradas. El botón queda resaltado en verde mientras está
  activo. Nunca se manda a ningún lado, solo se dibuja en ese celular.
- **Marcar entregado**: botón dentro del popup de cada cliente. El punto
  queda atenuado con una palomita (también se ve en los resultados del
  buscador), y arriba se cuenta cuántos van entregados. Se guarda solo
  en el celular de quien lo marca y se
  reinicia solo al otro día con la ruta nueva (no se sube a ningún repo).

## Historial de rutas

Tocando el texto **"Ruta del ..."** (arriba, junto al título) se abre un
calendario: los días con ruta guardada se ven normales y son tocables, los
que no tienen datos salen apagados. Al elegir un día aparece la etiqueta
**HISTÓRICO** y esa ruta (mismo mapa, colores, popups, Maps/Waze) — con un
botón **"Volver a hoy"** para salir. Mientras se ve un día del historial no
se pueden arrastrar puntos (eso solo aplica a "hoy"), pero "marcar
entregado" sigue funcionando igual, guardado por fecha.

Esto se archiva **solo**, sin ningún botón aparte: cada vez que se publica
una ruta, queda guardada **en ese mismo momento** en `historial/<fecha>.json`
dentro del repo privado (si se republica el mismo día, una corrección, se
reemplaza esa fecha). Por eso el histórico **no depende del maestro**: el
maestro se puede vaciar y volver a llenar cada día sin perder nada. Los
`MAPA Fxxx.xlsx` de cada día también quedan guardados en
`BASE\MAPAS\<fecha> <DÍA>`.

### Excel de control (histórico, ubicaciones y cambios de ruta)

`BASE\CONTROL DE RUTAS.xlsx` (botón **"Excel de control"** de la ventana; se
vuelve a armar solo después de cada publicación). Se arma únicamente con lo
ya publicado (el historial del repo privado), nunca con el maestro.

- **RUTAS**: cada cliente de cada ruta de los últimos 60 días, con filtros y
  un enlace "Ver" a Google Maps. **RESUMEN**: por fecha y ruta, de todo el
  histórico. **REVISAR UBICACIÓN**: clientes de la ruta vigente con
  ubicación sospechosa (lejos de su ruta, varios en el mismo punto,
  coordenada del sistema que cambió, o que no salen en el mapa).
- **UBICACIONES** (editable): todas las correcciones de ubicación (las mismas
  que el arrastre del modo admin). Se corrige/agrega pegando la coordenada
  de Google Maps en "Nueva coordenada", o se quita con Acción = QUITAR. Se
  aplican con "Aplicar ubicaciones" o al publicar. Si alguien movió ese
  punto en el mapa después de armar el Excel, no se pisa: queda en
  "Conflicto" para revisarlo.
- **CAMBIOS DE RUTA** (editable): pasar un cliente (o una ruta entera con
  TODOS) a otro chofer, **solo ese día** o **siempre**. Se aplica al generar
  la ruta: el cliente sale en el mapa y en el MAPA Excel del nuevo chofer
  (columna "Movido desde"), conserva su N° de liquidación original y no
  cuenta como un segundo cargue. La columna Estado dice si se aplicó o por
  qué no.

Es un archivo con datos de clientes: queda solo en esta PC (BASE\ no se
sube). Si se borra, se vuelve a armar y las reglas de CAMBIOS DE RUTA se
recuperan de `data\control_editables.json`.

### El maestro, liviano

`BASE\RUTAS DIARIAS CAMIONES - MY MAPS.xlsx` solo tiene la hoja SABANA DE
VENTA (sin tabla dinámica, "dia", RESUMEN ni modelo de datos). Cada día:
deja la fila 1 (encabezados), borra los datos de ayer desde la fila 2 y
pega los de hoy. **Ojo con el sábado**: si se trabaja sábado, pega juntos
los datos del viernes y del sábado, porque la ruta los une.

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

1. Pega la descarga del sistema en la hoja **SABANA DE VENTA** del maestro
   fijo **`BASE\RUTAS DIARIAS CAMIONES - MY MAPS.xlsx`** (siempre el mismo
   archivo; también se abre con el botón "Abrir en Excel" de la ventana) y
   **guarda**. Ya no hace falta actualizar la tabla dinámica de "dia", ni
   pegar en RESUMEN, ni separar por chofer a mano, ni crear la carpeta del día.
2. Doble clic en el acceso directo **"Actualizar Ruta"** del escritorio (o en
   `ACTUALIZAR RUTA.bat`). Se abre una ventana que:
   - Toma el maestro fijo de `BASE\`. Si no existiera, toma el archivo base
     más reciente de las carpetas del día en OneDrive
     (`MAPAS DIARIOS\<año>\<mes>\<día>`). Se puede usar otro con "Cambiar
     archivo...". Funciona aunque el Excel esté abierto.
   - Propone la fecha de liquidación más reciente. Si es **sábado**, marca
     también el **viernes** (se juntan). Se pueden marcar otras fechas.
   - Lista los camiones de esa fecha con sus clientes y total. Se tocan
     para desmarcar los que no van al mapa — eso se **recuerda** para la
     próxima vez. Los camiones nuevos salen en amarillo.
   - Si un camión tiene **2 cargues** (2 números de liquidación), sale en
     celeste con sus N° de liquidación; tocando su columna "Cargues" se
     elige **Unificado** (un solo mapa, como siempre) o **Separados**
     (`F325 CARGUE 1 - Liq 238` y `F325 CARGUE 2 - Liq 239`, numerados por
     N° de liquidación de menor a mayor, cada uno con su Excel y su ruta en
     el mapa).
   - **Todos los clientes deben tener ubicación (georreferencia).** Si a
     algún cliente de los camiones marcados le falta (sin coordenadas, 0,0 o
     fuera de Ecuador, y sin corrección), sale primero una **lista en rojo**
     arriba de los camiones (transportista, código, cliente, dirección,
     vendedor) y **"Generar y publicar" queda bloqueado** hasta corregirlos.
     Para corregir: elegir el cliente → **"1. Buscar en Google Maps"** (abre
     el mapa buscando el local) → clic derecho sobre el local → clic en los
     números (se copian) → **"2. Pegar ubicación copiada"** (se guarda al
     momento, igual que arrastrar el punto en el mapa). Para varios a la vez:
     "Varios: Excel de control" los pone arriba en UBICACIONES (filas en
     rojo); se pega la coordenada, se guarda y la lista se actualiza sola.
     "Copiar lista para vendedores" la deja lista para WhatsApp, agrupada por
     vendedor. Las correcciones se revisan contra lo vigente en GitHub
     (incluidas las hechas arrastrando en el mapa); sin internet se usa lo
     último descargado y se avisa.
3. **"Generar y publicar en el mapa"**: crea los `MAPA Fxxx.xlsx`
   (idénticos a los que salían de la tabla dinámica) en
   `BASE\MAPAS\<fecha> <DÍA>` — o junto al archivo base si se usó otro — y
   publica la ruta (archiva la anterior en el historial, igual
   que `publicar_datos.bat`). "Solo generar Excel" hace lo primero sin
   publicar.
4. Los choferes abren siempre el mismo link (arriba) y ven la ruta
   actualizada, sin instalar ni configurar nada (salvo pegar el código de
   acceso la primera vez).
5. Al terminar de publicar se abre sola la ventana **My Maps** (ver abajo):
   revisar la fecha de entrega y darle a **"Subir a My Maps"**.

### Google My Maps

Para los choferes que prefieren guiarse con Google Maps, la ruta del día
también se sube a My Maps, como se hacía a mano. Ventana **My Maps** (botón
"My Maps" de "Actualizar ruta", o doble clic en `MY MAPS.bat`):

- **Ruta del día**: una de las carpetas `BASE\MAPAS\<fecha> <DÍA>`.
- **Fecha de entrega**: va en el título. Se propone el siguiente día hábil
  (viernes y sábado → lunes); se cambia si ese día no se trabaja.
- Un mapa por grupo: **F1xx + F2xx**, **F3xx**, **FO** (otro prefijo, aparte).
  Si un grupo pasa de 10 camiones se parte en mapas parejos (My Maps admite
  10 capas por mapa); los cargues de un mismo camión van juntos.
- Cada mapa: título `1-MAPA DE ENTREGAS LUNES 12 DE OCTUBRE`, descripción
  `Transportistas F150-F151-...`, una capa por archivo (`MAPA F150.xlsx`)
  agrupada por **vendedor** con un color distinto por vendedor (los 30
  colores de My Maps, alternando para que no se confundan) y etiquetas con la
  **Razón**; todas las capas desmarcadas; compartido "cualquiera con el
  enlace" y **sin** búsqueda pública.
- Las ubicaciones van **corregidas** (las mismas del mapa web). Antes de
  abrir Chrome se revisa que **todos** los clientes tengan ubicación, con la
  misma regla del mapa web: si falta alguno, no se sube y se muestra la lista
  (hay que corregirlo en "Actualizar ruta" y volver a publicar). Se suben
  solo los MAPA que se publicaron (`_publicado.json` en la carpeta del día),
  no un Excel suelto de "Solo generar Excel".
- Abajo, los **links de los últimos 30 días** por fecha de entrega:
  "Copiar link", "Copiar mensaje del día" (título + transportistas + link de
  cada mapa, para WhatsApp) y "Abrir en el navegador". Al terminar de subir,
  el mensaje ya queda copiado.
- Los mapas con entrega de hace **más de 30 días** se mandan a la
  **papelera de Drive** en la siguiente subida (solo los que creó el robot;
  el mapa web guarda su historial completo). Si se vuelve a subir una
  entrega que ya estaba, los mapas anteriores van a la papelera (sus links
  dejan de servir).

Cómo funciona: My Maps no tiene API, así que `mymaps.py` maneja un **Chrome
aparte** (perfil propio en `data\chrome-mymaps`) con Playwright. La sesión
de Google se inicia **a mano** en ese Chrome (el robot nunca escribe
contraseñas; si Google pide verificar, se hace a mano y el robot espera
hasta 10 minutos). La cuenta va en `data\mymaps_config.json` y el registro
de mapas y links en `data\mymaps_registro.json`. Tarda ~50 s por capa: no
usar el Chrome del robot mientras trabaja. Si algo falla, el mapa a medias
va a la papelera, queda una captura en `data\mymaps_error.png` y se puede
volver a intentar (o subirlo a mano con los archivos de la carpeta del día).
Requiere `pip install playwright` (usa el Chrome instalado).

### Método anterior (sigue funcionando)

Con los `MAPA F1xx.xlsx` ya hechos en una carpeta: arrastra esa carpeta
sobre `actualizar_ruta.bat` (genera `data/latest.json`) y después corre
`publicar_datos.bat`. O a mano:
```
python convertir.py "C:\RUTAS DIARIAS\15MARTES 15"
```

## Estructura del proyecto

- `index.html`, `app.js`, `style.css` — el visor del mapa (selector de camión,
  mapa, pantalla de código de acceso).
- `control_rutas.py` — el Excel de control (histórico, ubicaciones, cambios
  de ruta); lo usa la ventana.
- `ACTUALIZAR RUTA.bat` + `generar_rutas.py` — la ventana del uso diario:
  lee la hoja SABANA DE VENTA, arma lo mismo que la tabla dinámica de "dia"
  (agrupa por Fecha Liq, Chofer, N° Liq, vendedor, cliente... y suma Peso,
  VentaNeta y Total), genera los MAPA y publica.
- `MY MAPS.bat` + `ventana_mymaps.py` — ventana de My Maps (subir y links);
  `mymaps.py` — el robot que crea los mapas en My Maps.
- `convertir.py` — convierte los `MAPA F1xx.xlsx` del día a `data/latest.json`.
- `data/` — solo en esta computadora (está en `.gitignore`, nunca se sube a
  este repo público): `latest.json` (datos del día), `generador_config.json`
  (último archivo base y camiones desmarcados) y `generador_error.log`.
- `actualizar_ruta.bat` — genera `data/latest.json` desde la carpeta del día.
- `publicar_datos.bat` — combina con GitHub, archiva la ruta anterior en
  `historial/` si cambió la fecha, y sube la nueva a `mapa-rutas-datos`.
- `archivar_dia.py` — lo usa `publicar_datos.bat`, no hace falta correrlo a mano.
- `TOKEN.md` — cómo generar y repartir el código de acceso.
- `sw.js` — service worker para que el mapa funcione sin conexión.

El repo privado `mapa-rutas-datos` es una carpeta aparte en
`C:\MAPA-RUTAS-DATOS` (clon de ese repo); `publicar_datos.bat` ya sabe copiar
ahí el archivo y subirlo. Ese mismo repo privado también guarda
`overrides.json` (correcciones de ubicación) y `historial/<fecha>.json`
(rutas de días anteriores) — nada de eso hay que tocarlo a mano.

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
