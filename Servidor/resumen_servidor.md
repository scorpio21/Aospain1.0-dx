# Resumen del Servidor - AoSpain

El servidor es el núcleo de la lógica del juego, manejando la persistencia de datos (archivos `.chr`), la IA de los NPCs y el cumplimiento de las reglas.

## Protocolo de Entrada (Cliente ➔ Servidor)
El servidor valida una "clave de cliente" (`orophin`) antes de permitir el login. Una vez logueado, procesa:
*   **Chat**: `;` (Hablar), `-` (Gritar), `\` (Susurrar), `/CLAN` (Mensaje de clan).
*   **Handshake**: `gIvEmEvAlcOde` (Protocolo de validación inicial).
*   **Habilidades**: `UK` seguido del índice de habilidad (Robar, Magia, Domar, Ocultarse).

## Comandos de Usuario
*   **Social/Clanes**: `/ONLINE`, `/ONLINECLAN`, `/FUNDARCLAN`, `/SALIRCLAN`.
*   **Estado**: `/EST` (Estadísticas), `/MUERTES` (Contador de frags), `/FAMA`.
*   **Interacción**: `/MEDITAR`, `/DESCANSAR`, `/COMERCIAR`, `/BOVEDA`, `/BALANCE`.
*   **NPCs Especiales**: `/RESUCITAR`, `/CURAR` (en sacerdotes), `/ENLISTAR`, `/RECOMPENSA` (facciones).

## Comandos de Administración (GMs)
El servidor distingue entre GMs rasos y "Dioses" (Privilegios >= 3):
*   **Gestión**: `/ECHAR`, `/BAN`, `/UNBAN`, `/INVISIBLE`, `/LIMPIAR` (Mundo).
*   **Teletransporte**: `/TELEP`, `/IRA` (Hacia un usuario), `/SUM` (Traer usuario), `/TELEPLOC`.
*   **Mundo/Spawn**: `/LLUVIA`, `/CC` (Lista de NPCs), `SPA` (Spawnear NPC), `/MATA` (Eliminar NPC).
*   **Dioses**: `/MOD` (Modifica ORO, EXP, NIVEL o BODY de cualquier usuario), `/DEST` (Destruye items), `/BLOQ` (Bloquea tiles), `/APAGAR` (Cierre del servidor).

## Seguridad y Cifrado
*   **Sistema Anti-Bot (Handshake):** Implementa un desafío inicial `gIvEmEvAlcOde`. El servidor genera un `ValCoDe` y un `RandKey` aleatorios que el cliente debe devolver procesados correctamente mediante la función `ValidarLoginMSG`.
*   **Validación de CRC:** Comprueba el CRC de cada paquete entrante. Si el CRC enviado por el cliente no coincide con el calculado por el servidor (`ServerSideCRC`), se cierra la conexión inmediatamente para prevenir herramientas de hackeo.
*   **Protección de Contraseñas:** Las contraseñas se gestionan con el algoritmo `MD5` para evitar filtraciones en caso de acceso no autorizado a los archivos `.chr`.
*   **Filtro de Cliente:** Solo permite conexiones si el cliente envía la cadena secreta de identificación (`orophin`).
*   **Logs de Seguridad:** Registra intentos de hackeo en `LogHackAttemp` cuando se detectan discrepancias en el protocolo o en el CRC.
