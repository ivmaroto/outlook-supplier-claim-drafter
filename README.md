# Generador de Borradores Outlook

Aplicación de escritorio en Python para preparar borradores de reclamación de pedidos pendientes a proveedores.

## Estado del proyecto

**Fase actual:** Fase 00 — Preparación del repositorio.

La aplicación funcional todavía no está implementada. Los requisitos, las decisiones técnicas y el estado de cada fase se documentarán conforme al protocolo maestro.

## Objetivo previsto

- Consultar información de pedidos pendientes en SQL Server en modo exclusivamente de lectura.
- Permitir revisar proveedores y líneas pendientes en una interfaz de escritorio.
- Preparar borradores de correo y archivos Excel para revisión del usuario.
- Mantener personalizaciones e historial local por usuario.
- Separar la configuración de desarrollo de la configuración corporativa.

Los mecanismos concretos de integración con Outlook moderno y las reglas de duplicidad del historial deberán investigarse y validarse antes de implementarse.

## Entorno de desarrollo

La versión propuesta inicialmente es Python 3.13 de 64 bits. Su compatibilidad con las dependencias y con el entorno corporativo está pendiente de validación.

La preparación del entorno virtual y las instrucciones de ejecución se documentarán cuando se hayan comprobado en el equipo de desarrollo.

## Estructura

- `src/`: código de la aplicación, pendiente de implementación.
- `tests/`: pruebas, pendientes de incorporación.
- `docs/`: protocolo, alcance, decisiones y estado de las fases.

## Seguridad

- No almacenar credenciales en el repositorio.
- No versionar bases de datos locales ni archivos de proveedores.
- No escribir datos en SQL Server.
- No enviar correos reales durante pruebas sin autorización explícita.
- Generar un borrador no equivale a enviar un correo.

## Protocolo de trabajo

Consultar `docs/protocolo-maestro.md` antes de realizar cambios.
