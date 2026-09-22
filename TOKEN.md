# Codigos de acceso (uno para choferes, uno para ti como admin)

Este paso lo tiene que hacer una persona con la cuenta de GitHub
`CoordinadorL` (la misma que ya se usa para este proyecto). No lo puede
hacer un asistente automatico por seguridad: es literalmente crear una
llave de acceso a datos privados.

Hay que generar **dos** tokens distintos:

| | Para quien | Permiso | Uso |
|---|---|---|---|
| **Codigo de chofer** | Todos los choferes (el mismo para todos) | Contents: **Read-only** | Solo ver el mapa |
| **Codigo de admin** | Solo tu | Contents: **Read and write** | Ver + arrastrar puntos para corregir ubicaciones |

**El codigo de admin es tuyo, no se comparte con los choferes.** Con ese
codigo cualquiera podria editar los datos, no solo verlos.

## Como generar cada uno

1. Entrar a: https://github.com/settings/tokens?type=beta
2. Click en **"Generate new token"**.
3. Llenar:
   - **Token name**: `mapa-rutas-lectura` (o `mapa-rutas-admin` para el tuyo)
   - **Expiration**: 90 dias o 1 año (cuando venza, se genera uno nuevo; es normal)
   - **Repository access**: **"Only select repositories"** → marcar solo
     `mapa-rutas-datos`. (Importante: NO darle acceso a otros repos.)
   - **Permissions** → **Repository permissions** → **Contents** →
     - Para el codigo de chofer: **Read-only**
     - Para tu codigo de admin: **Read and write**
     Todo lo demas se queda en "No access".
4. Click en **"Generate token"**.
5. GitHub lo muestra una sola vez (empieza como `github_pat_...`).
   **Copialo ya** — si cierras la pagina sin copiarlo, hay que generar otro.
6. Repite los pasos 2 a 5 para el segundo token.

## Como armar los links (para no tener que pegar el codigo a mano)

El codigo es largo, así que en vez de que cada chofer lo escriba o pegue,
mandas un link que ya lo trae incluido. Solo tienes que reemplazar
`TU_CODIGO` por el token real:

- **Para los choferes** (todos reciben el mismo link):
  ```
  https://coordinadorl.github.io/mapa-rutas-diarias/#codigo=TU_CODIGO_DE_CHOFER
  ```
- **Para ti** (guardalo tú, por ejemplo como marcador en tu celular/compu):
  ```
  https://coordinadorl.github.io/mapa-rutas-diarias/#admin=TU_CODIGO_DE_ADMIN
  ```

Cada uno abre su link **una sola vez** y el navegador guarda el codigo
correspondiente; después ya no hace falta el link, solo entrar al sitio
normal. Si prefieren escribir el codigo a mano en vez de usar un link,
también funciona (pantalla "Código de acceso" al abrir el sitio).

El codigo nunca queda en el repo ni en el codigo fuente del sitio: vive
solo en el link que mandas por WhatsApp y en el navegador de cada quien.

## Modo administrador: corregir la ubicacion de un cliente

Al entrar con tu link de admin ves una etiqueta **ADMIN** junto al título.
Los puntos se pueden arrastrar; al soltar uno, se guarda automáticamente
(abajo aparece "Guardando ubicacion..." y despues "Ubicacion guardada").
Esa corrección queda asociada al cliente (no al día), así que se mantiene
aunque el archivo de datos se regenere al día siguiente.

## Si se filtra o hay que cambiarlo

1. Ir a https://github.com/settings/tokens?type=beta y borrar (`Revoke`)
   el token viejo (el de chofer, el de admin, o ambos según el caso).
2. Generar uno nuevo siguiendo los pasos de arriba.
3. Armar el link nuevo y repartirlo de nuevo — todos van a tener que
   abrirlo una vez más.
