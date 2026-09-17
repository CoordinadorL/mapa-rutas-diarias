# Generar el codigo de acceso (una sola vez)

Este paso lo tiene que hacer una persona con la cuenta de GitHub
`CoordinadorL` (la misma que ya se usa para este proyecto). No lo puede
hacer un asistente automatico por seguridad: es literalmente crear una
llave de acceso a datos privados.

1. Entrar a: https://github.com/settings/tokens?type=beta
2. Click en **"Generate new token"**.
3. Llenar:
   - **Token name**: `mapa-rutas-lectura`
   - **Expiration**: 90 dias o 1 año (cuando venza, se genera uno nuevo y se
     reparte de nuevo a los choferes; es normal).
   - **Repository access**: elegir **"Only select repositories"** y marcar
     solo `mapa-rutas-datos`. (Importante: NO darle acceso a otros repos.)
   - **Permissions** → **Repository permissions** → buscar **Contents** →
     ponerlo en **Read-only**. Todo lo demas se queda en "No access".
4. Click en **"Generate token"**.
5. GitHub muestra el codigo una sola vez (empieza como `github_pat_...`).
   **Copialo y guardalo ya** (en una nota, en el gestor de contraseñas, etc.)
   — si se cierra la pagina sin copiarlo, hay que generar uno nuevo.

## Como se reparte

Ese codigo es la "llave" para ver los datos de clientes. Se comparte por un
medio privado (WhatsApp del grupo de choferes, por ejemplo), **nunca** se
sube a GitHub ni se pega en ningun archivo del repo publico.

Cada chofer lo pega **una sola vez** en el mapa (se lo va a pedir la primera
vez que lo abren en su celular) y despues queda guardado en ese telefono.

## Si se filtra o hay que cambiarlo

1. Ir a https://github.com/settings/tokens?type=beta y borrar
   (`Revoke`) el token viejo.
2. Generar uno nuevo siguiendo los pasos de arriba.
3. Repartir el nuevo codigo a los choferes (van a tener que pegarlo de
   nuevo, una sola vez).
