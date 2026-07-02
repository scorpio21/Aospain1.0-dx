# Propuestas de Mejora del Servidor - AoSpain

El servidor puede beneficiarse de cambios estructurales para soportar más usuarios y ser más difícil de vulnerar:

## 1. Persistencia de Datos (Base de Datos)
*   **Migración a SQL:** Actualmente el servidor usa archivos `.chr` (tipo INI). Esto es ineficiente para búsquedas y escalabilidad. Migrar a **SQLite** o **MySQL** permitiría consultas rápidas, copias de seguridad en caliente y una gestión de usuarios mucho más robusta.

## 2. Rendimiento y Estabilidad
*   **Manejo de Buffers (ByteQueue):** Reemplazar el procesamiento de strings en `HandleData` por un sistema de colas de bytes. Esto evita la fragmentación de memoria y mejora la velocidad de respuesta.
*   **Multithreading para Tareas Pesadas:** Mover tareas como el "WorldSave" o el procesamiento de logs a hilos separados (o procesos externos) para evitar el "congelamiento" momentáneo del juego.

## 3. Seguridad y Anti-Cheat
*   **Sistema de Logueo Robusto:** Reemplazar el intercambio simple de `ValCoDe` por un sistema de intercambio de claves Diffie-Hellman para cifrar la sesión de forma única.
*   **Validaciones en Lado Servidor:** Asegurar que todas las acciones (distancia de ataque, velocidad de movimiento, uso de hechizos) se validen estrictamente en el servidor para anular el uso de "Speedhacks" o "Rangehacks".

## 4. Nuevos Sistemas
*   **Sistema de Eventos Automáticos:** Refactorizar la lógica de torneos para que sea totalmente configurable desde un archivo externo, permitiendo programar eventos sin reiniciar el servidor.
*   **Sistema de Logs Centralizado:** Implementar un sistema de logs en base de datos para rastrear intercambios de items sospechosos (anti-dupe) de forma más sencilla.
