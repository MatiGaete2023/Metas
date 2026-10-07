# Integración del Descargador SITFA — revisión 2

Fuente estudiada: `02_FUNCIONAMIENTO_INTERNO_Y_FLUJO_COMPLETO.md`, versión de referencia SITFA Descargador 2.6.0 corregida V4. Este documento permite completar el diseño sin subir el ejecutable completo. Las afirmaciones sobre su funcionamiento proceden de esa referencia; no se ha ejecutado ni inspeccionado el descargador aquí.

## 1. Decisión de integración

Conservar la arquitectura existente: **aplicación Windows + extensión Chrome + sesión SITFA iniciada por el usuario**. Incorporar al coordinador local los motores de consolidación y Meta 4 y ampliar su planificador para varias modalidades y combinaciones específicas.

No trasladar la autenticación a la aplicación. No sustituir las consultas por una automatización nueva del navegador. La extensión ya resuelve aspectos delicados del sistema: `Envio()`, estabilización del formulario, construcción del POST, descarga autenticada y reconstrucción de evidencia.

El destino deseado es una sola aplicación local con la extensión como componente necesario. Tener una herramienta integrada no significa eliminar la extensión. La distribución seguirá necesitando recursos y dependencias; el `.exe` no se copiará aisladamente.

El código exacto del ejecutable original no está completo según la propia referencia. Por eso no se promete reutilización directa de sus clases internas. Si los módulos necesarios están disponibles, se integran; si solo existe el binario, hay que verificar una interfaz soportada antes de coordinarlo externamente. Conocer `/next` no equivale a conocer todo el protocolo ni a disponer de una API de control del ejecutable.

## 2. Flujo integrado revisado

```text
Usuario inicia sesión en SITFA en Chrome
             ↓
RUS Integral inicia puente local y muestra código temporal
             ↓
Extensión vinculada desde la pestaña institucional
             ↓
catalog → opciones vigentes del formulario
             ↓
Plan: tribunales + modalidades + pestañas + filtros
             ↓
Tareas individuales → prepare → Envio() → reaplicar → POST final
             ↓
Consulta en sesión autenticada
       ├── Resultados → Excel institucional original
       ├── Cero → evidencia PDF vinculada a la consulta
       └── Fallo/rechazo/sesión vencida → incidencia, nunca cero
             ↓
Manifiesto y validación de cobertura
             ↓
Consolidado Espera/Cumplimiento → revisión de FECHA OBS.
             ↓
Meta 4 y controles
```

La ruta **mi consolidado → Meta 4** omite completamente conexión, catálogo y descarga. Puede utilizarse sin Chrome ni sesión SITFA.

El documento usa SITFA y `familia.pjud.cl`, mientras el usuario llama RUS al origen. En los contratos internos se identifica la pantalla SITFA concreta; el nombre de producto RUS Integral se conserva provisionalmente.

## 3. Qué ya existe y qué se amplía

| Capacidad | Documentada en 2.6.0 V4 | Trabajo del producto integrado |
|---|---|---|
| Sesión autenticada en Chrome | Sí, login manual | Conservar y mostrar estado de vinculación |
| Puente local con código temporal | Sí | Conservar contrato y desconexión; evitar registrar la clave en productos |
| Catálogo leído de SITFA | Sí | Usar valores/códigos reales para cada pantalla |
| Varios tribunales y pestañas | Sí | Integrarlos al plan común |
| Varias modalidades por lote | No queda documentado; se menciona una modalidad | Expandir el plan a tareas con distinta modalidad |
| Cruces arbitrarios | No acreditados | Tabla de combinaciones individuales seleccionables |
| Pausa/continuar/cancelar | Sí | Conservar; cancelar no borra archivos válidos |
| Reanudación tras reiniciar | No acreditada | Persistencia del plan y comprobación de tareas; nueva vinculación al reiniciar |
| Excel institucional y PDF de cero | Sí | Relacionarlos con tareas y validación del consolidado |
| Manifiesto persistente para consumidores | No consta su formato ni existencia | Añadirlo como contrato explícito de intercambio |
| Consolidación/Meta 4 | No descritas como parte del descargador | Integrar motores estudiados y nueva entrada desde Excel propio |

La lista de pantallas reconocidas no demuestra que cada una soporte todas las operaciones o filtros. Órdenes de búsqueda no aparece descrito como un flujo específico en la referencia: se mantiene como necesidad del usuario y se verifica su correspondencia con una pantalla/acción antes de ofrecerla.

## 4. Reglas V4 que se deben conservar

1. Catálogos de los controles actuales, sin listas institucionales inventadas en la aplicación.
2. Inputs RIT, RUT y dígito verificador no utilizados quedan vacíos; cero no sustituye una cadena vacía. Selects usan una opción neutral válida.
3. Configurar → `Envio()` → verificar rechazo → reaplicar valores → limpiar campos no usados → construir POST. Este orden forma parte del comportamiento requerido.
4. Validar campos permitidos, duplicados inesperados e `irAccion` por pantalla.
5. Los pares del **POST efectivamente enviado** son la fuente para la evidencia y el manifiesto; los filtros deseados se conservan aparte para comparar que coincidan.
6. Mantener sesión en el navegador, `credentials: "include"`, `cache: "no-store"`, timeout, límites de tamaño y rechazo de respuestas inesperadas.
7. Validar enlaces institucionales con las restricciones documentadas. No ampliar hosts para adaptar la consolidación.
8. Conservar bytes del Excel descargado; la normalización se realiza después y en otros archivos.
9. Restaurar la pantalla original al terminar o cancelar mediante `unlock`, registrando si no se pudo completar la restauración.
10. Rotar el código al reiniciar/desconectar y no exportar cookies, claves temporales ni encabezados de conexión en informes.

Los detalles de encabezados, serialización, retorno de resultados y versiones del protocolo aún deben verificarse. No se inventan nombres de endpoints adicionales a `/next`.

## 5. Contrato de intercambio propuesto

El coordinador necesita que cada resultado se vincule inequívocamente a una tarea. Los siguientes campos son una **propuesta nueva**, no el formato confirmado del ejecutable:

```text
Run
  run_id, downloader_version, plan_version, started_at

Task
  task_id, run_id, screen, tribunal_code, modality_code, tab
  requested_filters, requested_period, expected_action

Result
  task_id, attempt_id, status, queried_at
  effective_post_pairs, detected_record_count
  institutional_report_date, errors

Artifact
  task_id, attempt_id, kind, relative_path, byte_length, sha256

ZeroEvidence
  task_id, attempt_id, zero_observed, capture_time
  effective_post_hash, reconstructed_dom_hash, pdf_artifact
  validation_status, validation_issues
```

Los valores efectivos del POST pueden contener datos de identificación en otros ítems; se conservan localmente y el resumen exportable evita exponerlos sin necesidad. Las fechas de consulta, emisión del reporte, emisión Meta 4 y FECHA OBS. son campos separados.

Una salida documental verificable puede ser un `manifiesto.json` por lote más un registro persistente local. Para mover una carpeta se usan rutas relativas y huellas; una coincidencia de nombre no basta para enlazar archivo y tarea.

Al reanudar tras reiniciar, vincular de nuevo la extensión y verificar fuentes del mismo plan/corte. Un catálogo cambiado, filtros distintos o un corte nuevo requieren nuevas tareas. Pausar un lote no permite hacer pasar archivos de otro corte como una extracción vigente.

## 6. Cobertura y PDF de cero

El estudio anterior detectó correctamente una limitación del motor Meta 4: interpreta PDF/imágenes por su nombre y ruta, sin leerlos. La nueva referencia explica que algunos PDF tienen una función legítima de evidencia de cero. Ambos hechos son compatibles.

Se distinguen:

| Entrada | Tratamiento propuesto |
|---|---|
| Excel legible con filas | Cubrir combinación y aportar registros |
| Excel legible vacío con estructura y filtros comprobados | Cero confirmado según validación |
| PDF generado en el flujo de evidencia y enlazado a tarea/POST/resultado | Puede acreditar cero tras controles del lote |
| PDF suelto, sin manifiesto o asociación verificable | Evidencia pendiente; revisión manual documentada |
| HTML de login, timeout, respuesta vacía, error o consulta rechazada | Fallo; no acreditar cero |

Para acreditar cero con el nuevo manifiesto se comprueba: tarea correcta, POST efectivo equivalente al plan, resultado reconocido como cero, captura asociada y archivo íntegro. El consumidor del consolidado no necesita OCR para un PDF generado y trazado por el propio proceso; los PDF externos son un caso distinto.

El SHA-256 del DOM asegura consistencia entre la reconstrucción validada y la vista mostrada dentro del flujo descrito. No prueba por sí mismo que el servidor devolvió cero o que el PDF provino del descargador. La huella del PDF es un campo adicional propuesto, distinto del hash del DOM ya documentado.

El flujo de evidencia ejecuta nuevamente el POST en un iframe. La implementación deberá comprobar también el resultado de esa segunda consulta: si cambia respecto de la primera o deja de ser cero, no se certificará el vacío y se repetirá/revisará la tarea. Restaurar filtros en el DOM describe la consulta; no modifica ni certifica por sí solo los datos devueltos.

Conflicto importante: Excel con filas y PDF de cero para la misma combinación. Comparar tarea, filtros y corte de ambos. No combinar momentos distintos ni sustituir automáticamente el Excel por cero; el usuario resuelve cuál corresponde al producto. Una fecha de modificación del archivo no acredita vigencia institucional.

## 7. Cambios en la interfaz

La conexión añade una pantalla sencilla:

- Abrir SITFA e iniciar sesión en Chrome.
- Vincular la extensión con el código temporal que muestra la aplicación.
- Mostrar conexión, pantalla detectada y botón **Cargar opciones**.
- Mantener indicación de las dos pestañas necesarias y del estado del lote.

El código aparece solo durante vinculación; los detalles del POST y hashes quedan en controles técnicos, no en el flujo principal. Los usuarios eligen tribunales/modalidades/filtros con sus etiquetas institucionales.

Las vistas de cobertura ofrecen **Ver evidencia** para ceros confirmados y **Revisar evidencia** cuando falta trazabilidad. La maqueta revisada muestra este escenario con datos ficticios; no establece una conexión real.

CSMP Integral se utilizará únicamente como referencia visual. No se deriva de él ningún requisito funcional.

## 8. Verificación de integración futura

Conservar las diez regresiones recomendadas en la fuente: Espera y Cumplimiento con resultados, rango de fechas, PDF de cero con filtros correctos, inputs vacíos, persistencia de fechas/tribunal/modalidad tras `Envio()`, Excel legible, cancelar sin borrar archivos y restauración original.

Añadir comprobaciones del producto integrado:

1. Dos tribunales y dos modalidades generan los cruces elegidos; modo específico ejecuta solo las filas marcadas.
2. Catálogo y valores de cada tarea coinciden con el formulario detectado.
3. Sesión vencida y rechazo institucional no producen cero confirmado.
4. La evidencia usa el POST final; cambios de resultado en la repetición quedan pendientes.
5. Cada archivo guardado queda asociado a una tarea y su huella.
6. Cancelar conserva artefactos y registra qué tareas no terminaron.
7. Recuperar tras cierre exige nueva conexión y verifica el plan/corte antes de reutilizar artefactos.
8. Excel/PDF contradictorios quedan como conflicto, sin pérdida silenciosa de filas.
9. Consolidados de cualquier modalidad mantienen sus registros; Meta 4 informa las filas fuera de su perfil.
10. Importar consolidado propio conserva FECHA OBS. y no toca el archivo, sin necesidad de sesión ni extensión.

Estas verificaciones son criterios para la implementación, no resultados obtenidos en esta entrega.

## 9. Qué permite avanzar sin subir el ejecutable

El documento basta para completar la arquitectura y diseñar el contrato de integración. También permite implementar por separado consolidación flexible y Meta 4 desde consolidado propio, que no dependen del descargador.

Para la fase de integración real puede trabajarse sobre un repositorio accesible o los módulos fuente concretos que estén disponibles, especialmente coordinador local y archivos de la extensión. No hace falta pedir otra vez toda la distribución portable. Si solo existe el ejecutable, se comprueban sus interfaces en un entorno Windows donde esté instalado antes de decidir el método de integración.

Los trabajos anteriores mencionados por el usuario no están disponibles automáticamente como código en este espacio. Esta revisión se fundamenta en los adjuntos y la documentación efectivamente recibidos.
