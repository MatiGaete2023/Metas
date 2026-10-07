# Plan de trabajo

Este repositorio es el destino autorizado por el usuario para guardar avances y materiales necesarios del proyecto.

## Completado

- Estudiar los motores Meta 4 v2.4 y Consolidador Excel v1.2.
- Comprobar las 57 pruebas de Meta 4 y el traslado de FECHA OBS. con datos ficticios.
- Diseñar descarga, consolidación flexible y ambas entradas de Meta 4.
- Analizar documentalmente el Descargador SITFA 2.6.0 corregida V4.
- Preparar y comprobar una maqueta interactiva con conexión simulada, combinaciones y fechas.
- Conservar fuentes originales y documentación en Git.

## Siguiente desarrollo

1. Extraer un modelo e importadores comunes, separando reportes institucionales de consolidados propios.
2. Implementar consolidación flexible por filas y combinaciones, con trazabilidad y controles separados.
3. Implementar la entrada consolidado propio → Meta 4, conservando FECHA OBS. y sin sobrescribir el archivo de entrada.
4. Verificar los contratos disponibles del coordinador/extension del descargador e integrar sus consultas, catálogo y evidencia.
5. Ampliar el planificador a varias modalidades y cruces específicos; añadir persistencia y reanudación verificada.
6. Ajustar la apariencia con la referencia visual de CSMP Integral.
7. Validar formato institucional y probar el producto completo en Windows/SITFA.

## Referencias pendientes

- Apariencia visual de CSMP Integral; no aporta requisitos funcionales.
- Plantilla institucional Meta 4 validada y muestras anonimizadas para comparación.
- Contratos o módulos disponibles del descargador para verificar la integración. El documento técnico ya permite continuar el diseño sin subir el ejecutable completo.

El análisis y los criterios de aceptación detallados están en [el estudio](diseno/Estudio_y_diseno_RUS_Integral.md) y [el anexo SITFA](diseno/Anexo_integracion_Descargador_SITFA.md).
