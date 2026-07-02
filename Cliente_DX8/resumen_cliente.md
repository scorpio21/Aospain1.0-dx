# Resumen del Cliente - AoSpain

El cliente gestiona la interfaz gráfica (DirectX 7), el sonido (módulo Bass) y la comunicación bidireccional con el servidor a través de paquetes de texto.

## Protocolo de Entrada (Servidor ➔ Cliente)
El procedimiento `HandleData` en `TCP.bas` procesa los paquetes entrantes:
*   **`LOGGED`**: Confirmación de inicio de sesión exitoso.
*   **`PU`**: Actualización de la posición del usuario en el mapa (`PUx,y`).
*   **`ERR`**: Muestra un mensaje de error al usuario.
*   **`||`**: Mensaje de consola (chat, información del sistema).
*   **`INITCOM` / `FINCOMOK`**: Inicia o finaliza el modo de comercio con NPCs.
*   **`CFX`**: Crea un efecto visual (FX) en un personaje específico.
*   **`NAVEG`**: Alterna el modo navegación (barcos).

## Comandos y Acciones de Salida (Cliente ➔ Servidor)
*   **Movimiento**: Envía `M1` (Norte), `M2` (Sur), `M3` (Este), `M4` (Oeste).
*   **Interacción**: `AT` (Atacar), `AG` (Agarrar item), `USA` (Usar item del inventario).
*   **Combate**: `TAB` (Modo combate), `SEG` (Seguro de ataque).
*   **Clics**: `LCx,y` (Clic izquierdo/Mirar), `RCx,y` (Clic derecho/Acción).

## Seguridad y Cifrado
*   **Integridad de Paquetes (CRC):** Utiliza `PrivateCrcFunction.bas` para generar un código de verificación por cada paquete enviado. El CRC es dinámico y depende del paquete anterior, lo que dificulta la creación de bots.
*   **Cifrado de Mensajes:** El módulo `Mod_Cripto.bas` se encarga de ofuscar las cadenas de datos sensibles antes de que salgan del cliente.
*   **Validación de Identidad:** Envía una clave de aplicación fija (`orophin`) y resuelve desafíos numéricos (`ValCoDe`) enviados por el servidor para demostrar que es el cliente oficial.
*   **Hashing:** Emplea `MD5.bas` para el manejo seguro de contraseñas localmente antes de ser enviadas.
