# Alcance del proyecto

## Objetivo

Preparar borradores de reclamación de pedidos pendientes a proveedores mediante una aplicación de escritorio.

## Requisitos funcionales confirmados

- Interfaz de escritorio Python con Tkinter/ttk.
- Consulta de SQL Server exclusivamente en modo de lectura.
- Tratamiento de datos con pandas y pyodbc.
- Persistencia local con SQLite.
- Generación de Excel con openpyxl.
- Preparación de borradores para Outlook moderno, con mecanismo pendiente de investigación y validación.
- Configuración separada para desarrollo y entorno corporativo.
- Personalización local de asunto y cuerpo por proveedor.
- Selección individual de líneas pendientes y proveedores para generación.
- Historial local para ayudar a evitar reclamaciones repetidas.

## Fuera de alcance de la Fase 00

- Implementación de la interfaz funcional.
- Conexión a SQL Server.
- Integración con Outlook.
- Generación de Excel funcional.
- Definición de las reglas de duplicidad del historial.
- Envío automático de correos.

## Pendiente de validación

- Esquema real de SQL Server y consultas autorizadas.
- Compatibilidad de las dependencias con la versión de Python elegida.
- Mecanismo viable para crear borradores en Outlook moderno.
- Reglas exactas para duplicidad y estados del historial.
- Configuración concreta del entorno corporativo.
