# Validación del estudio

Comprobaciones realizadas el 7 de octubre de 2026. Solo se utilizaron datos ficticios y archivos temporales. No se modificaron los motores originales.

## Motor Meta 4 recibido

Se ejecutó `python -m pytest tests -q` en el paquete v2.4: **57 pruebas aprobadas en 33,04 segundos**. Estas pruebas verifican clasificación, formatos Excel, cobertura, duplicados, conflictos y comportamiento del CLI, entre otros aspectos incluidos en la suite.

Entorno: Python 3.12; pandas/openpyxl/xlrd disponibles, pytest 9.1.1 y xlwt 1.3.0 instalados para esta comprobación. Holidays no estaba instalado, por lo que se usó el calendario de respaldo del paquete. No se probaron los BAT en cmd.exe ni una sesión real de RUS.

## Comprobaciones adicionales con datos ficticios

| Comprobación | Resultado |
|---|---|
| Excel propio con hojas Espera y Cumplimiento y columna `fecha obs` | Ambas hojas leídas por `read_registros()` |
| Traslado de fecha de observación 07-10-2026 a Meta 4 | Fecha presente en la octava columna de ambas hojas FAE pertinentes |
| Archivo de entrada después de la generación por llamadas directas | Contenido binario sin cambios |
| Consolidador con tribunal fuera del catálogo Meta 4 y modalidad ambulatoria | Conserva el registro; no aplica restricción FAE/Residencia |
| Columna adicional FECHA OBS. en reporte del consolidador | Descartada por la proyección a sus 21 columnas; confirma la necesidad de corregirlo |
| Título Cumplimiento con columnas de Espera | El consolidador rechaza la contradicción |

El traslado de fechas del motor Meta 4 se comprobó invocando sus funciones directamente. Su CLI actual no ofrece la ruta independiente desde un consolidado propio.

## Maqueta interactiva

Comprobada con Playwright y Chromium del sistema:

- El perfil Meta 4 genera 68 combinaciones.
- Retirar un cruce en modo específico reduce la selección a 67.
- Cambiar FECHA OBS. de DEMO-01 a 07-10-2026 se refleja en la vista FAE Espera.
- Vaciar las fechas conserva celdas vacías en la vista final.
- La pantalla de inicio cabe en una ventana móvil de 390 px sin desbordamiento horizontal.
- No se registraron errores JavaScript durante ese recorrido.

Se entregan capturas de inicio y vista previa. La maqueta no importa Excel, no genera Meta 4 real y no se conecta a RUS. Sirve para revisar navegación, selección y transferencia visual de fechas.

## Pendientes de verificación real

La revisión 2 incorpora y analiza el documento técnico del Descargador SITFA 2.6.0 corregida V4. No se ha ejecutado ni inspeccionado su código: su arquitectura se conoce documentalmente. Continúan pendientes la referencia visual CSMP Integral, reportes reales anonimizados, plantilla institucional validada y comprobación de los contratos del descargador. Las pruebas aprobadas no acreditan esos componentes ni el funcionamiento del producto integrado, que se entrega en fase de diseño.

## Alcance de la revisión 2

Las 57 pruebas y comprobaciones de fechas corresponden al estudio anterior y no se han repetido porque los motores no cambiaron. La revisión 2 añade análisis documental y modifica el diseño y la maqueta; no constituye una prueba de integración con SITFA.

La maqueta de revisión 2 se comprobó con Playwright/Chromium: entrada a conexión desde Descargar; catálogo de ejemplo deshabilitado antes de simular vinculación y habilitado después; perfil de 68 cruces y selección específica de 67; cero confirmado ficticio en cobertura; traslado de FECHA OBS. a FAE Espera; inicio móvil sin desbordamiento y ausencia de errores JavaScript en ese recorrido. Se actualizaron las capturas y se añadieron conexión y cobertura. No se probó una conexión institucional.
