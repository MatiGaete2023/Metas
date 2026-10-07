# RUS Integral — estudio y diseño funcional

Fecha: 7 de octubre de 2026. Revisión 2: incorpora el documento técnico del Descargador SITFA 2.6.0 corregida V4. Alcance de esta entrega: análisis de los dos paquetes recibidos y de la documentación del descargador, especificación de la herramienta integrada y maqueta interactiva de sus flujos. La maqueta usa datos ficticios; no se conecta a RUS ni procesa archivos reales.

## 1. Resultado del estudio

La integración es viable. Conviene conservar los lectores Excel y el generador de formato Meta 4, pero separar descarga, normalización, consolidación y generación en módulos independientes. Encadenar los BAT actuales conservaría sus limitaciones: cobertura fija, reglas distintas de errores y sobrescritura del consolidado al regenerar Meta 4.

El cambio central consiste en convertir el consolidado en un producto editable y reutilizable. El usuario podrá completar `FECHA OBS.` en Excel o en la herramienta y generar Meta 4 desde esa versión, sin volver a descargar ni reconstruir los registros.

### Material efectivamente recibido

| Paquete | Contenido relevante | Situación |
|---|---|---|
| META4_Automatizado_v2_4.zip | Motor Python, lanzadores Windows, calendario adicional, documentación y pruebas | Analizado por lectura de código y pruebas |
| Consolidador_Excel_Espera_Cumplimiento.zip | Motor Python, selector gráfico de carpeta, lanzadores y documentación | Analizado por lectura y comprobación con datos ficticios |
| Descargador SITFA 2.6.0 corregida V4 | Documento 02_FUNCIONAMIENTO_INTERNO_Y_FLUJO_COMPLETO.md | Arquitectura y flujo estudiados documentalmente; código y ejecutable no inspeccionados |
| CSMP Integral | No hay código, capturas ni recursos visuales en los ZIP | Apariencia exacta pendiente de referencia |

No se ha comprobado acceso a RUS, compatibilidad con sus pantallas actuales, ejecución de los BAT en Windows ni resultados con reportes reales. Tampoco se ha aportado una plantilla oficial externa de Meta 4: el paquete genera su propia plantilla cuando no encuentra `assets/meta4-template.xlsx`, archivo que no está incluido en el ZIP.

## 2. Qué hace cada herramienta y cómo lo hace

### 2.1. Meta 4 automatizado, versión 2.4

Entrada: carpeta y subcarpetas, mes o fecha de emisión y plantilla opcional. Python con pandas, openpyxl y xlrd; holidays es opcional. Los BAT detectan Python 3.10 o superior y lanzan el script. El selector intenta Tkinter y, en Windows, PowerShell.

Secuencia real:

1. Inventaría Excel, PDF, imágenes y archivos ignorados. Excluye sus salidas conocidas y temporales de Excel.
2. Lee los Excel y busca, entre las primeras 30 filas, encabezados RIT, TRIBUNAL y DERIVACIÓN. Si encuentra varias hojas candidatas, escoge una por cantidad de registros y encabezados.
3. Determina Espera/Cumplimiento por A2, encabezados y, como respaldo, ruta. Si título y columnas se contradicen, prevalece el título y deja una advertencia.
4. Determina tribunal con la columna TRIBUNAL, A1 y ruta. Tiene catálogo de 17 tribunales y alias. Los archivos con varios tribunales se separan por tribunal de cada fila.
5. Clasifica derivaciones: FAE si comienzan con FAE; Residencia si comienzan con R o CREAD. Las demás filas se excluyen del consolidado y Meta 4, con advertencias.
6. Crea candidatos por tribunal, modalidad y estado. Espera 17 × 2 × 2 = 68 combinaciones.
7. Escoge un candidato por combinación. Excel prevalece sobre PDF o imagen; luego decide la fecha de emisión o de modificación y otros desempates. No fusiona automáticamente varios Excel de una misma combinación.
8. Escribe `Registros.xlsx`, lo vuelve a leer y reconstruye las cuatro hojas Meta 4 por bloques de tribunal.
9. Verifica conteos, totales y filas idénticas; genera reporte TXT. Faltantes y numerosos errores no impiden generar archivos.

Productos:

| Producto | Contenido |
|---|---|
| Registros.xlsx | ESPERA, CUMPLIMIENTO, INVENTARIO_ARCHIVOS, MATRIZ_COBERTURA, ERRORES_Y_ADVERTENCIAS |
| META 4 - actualizado AAAA-MM.xlsx | Residencia Espera, Residencia Cumplimiento, FAE Espera, FAE Cumplimiento |
| Reporte TXT | Inventario, cobertura, selecciones, incidencias y totales |

Las hojas de registros ya incluyen `FECHA OBS.`, `ARCHIVO ORIGEN` y `RUTA ORIGEN`. `read_registros()` lee la fecha de observación y la coloca en la octava columna de Meta 4. Sin embargo, `main()` siempre escanea la carpeta y escribe nuevamente `Registros.xlsx` antes de leerlo. No existe una opción CLI para generar exclusivamente desde un consolidado propio. Esa es la causa técnica del problema señalado por el usuario.

El informe final contiene ocho columnas. En Espera: RIT, TRIBUNAL, TIPO, SEXO, DERIVACIÓN, T ESPERA, FEC. RESOLUCIÓN, FECHA OBS. En Cumplimiento reemplaza T ESPERA por FEC. RESOLUCIÓN y utiliza FEC. INGRESO EFECTIVO en la séptima columna. No exporta RUT y nombre en esas hojas finales.

La fecha de emisión predeterminada es el primer día hábil del mes, con calendario chileno y feriados adicionales; se puede indicar una fecha exacta. Es una fecha distinta de `FECHA OBS.`.

### 2.2. Consolidador Excel, versión 1.2

Entrada: carpeta seleccionada con navegador gráfico Tkinter o ruta por CLI. Lee recursivamente .xls, .xlsx y .xlsm mediante xlrd y openpyxl. Copia valores; en .xlsx/.xlsm usa `data_only=True`, por lo que depende del valor guardado por Excel si una celda contiene una fórmula.

Secuencia real:

1. Excluye temporales, consolidaciones anteriores con su nombre conocido y carpetas del etiquetador anterior.
2. Selecciona una hoja por archivo, buscando RIT, TRIBUNAL y DERIVACIÓN en las primeras 30 filas y ordenando por registros y encabezados.
3. Identifica Espera/Cumplimiento por título de A2 y columnas. Una contradicción genera error.
4. Exige todos los encabezados de su esquema de 21 columnas para el estado detectado.
5. Copia las filas con RIT no vacío. No restringe tribunal ni distingue FAE/Residencia. No tiene selector de modalidades o combinaciones.
6. Agrupa coincidencias por RUT normalizado + nombre normalizado entre ambos estados y las muestra en Duplicados. No elimina esas filas.
7. Si cualquier archivo falla, no genera un consolidado parcial. Si todo es válido, crea una salida numerada sin sobrescribir.

Salida: `Registros consolidados.xlsx`, con Espera, Cumplimiento y Duplicados. Encabezados coloreados, filtros, panel congelado, anchos y fechas formateadas. No conserva `FECHA OBS.` ni origen en las hojas principales: proyecta únicamente sus encabezados fijos.

### 2.3. Descargador SITFA 2.6.0 corregida V4

Fuente: documento técnico aportado por el usuario, incorporado sin cambios como `Fuente_Descargador_SITFA_2_6_0_V4.md`. Lo siguiente describe el funcionamiento documentado, no una ejecución comprobada aquí. El usuario denomina RUS al origen del trabajo; la referencia identifica SITFA y el host `familia.pjud.cl`. Mantendremos esa distinción técnica hasta confirmar la relación entre las pantallas utilizadas.

La herramienta es portable para Windows x64. Contiene `SITFA_Descargador.exe`, empaquetado con PyInstaller y Python 3.12, una carpeta `_internal` con sus dependencias y una extensión Chrome Manifest V3. El ejecutable necesita la distribución completa.

La aplicación inicia un puente local en `127.0.0.1` y muestra un código temporal `PUERTO-CLAVE`. El usuario abre una sesión SITFA en Chrome, abre desde esa pestaña la página de conexión de la extensión y vincula ambos componentes con el código. No automatiza el login ni recibe la contraseña o las cookies de SITFA.

La extensión consulta `/next` para obtener tareas. Ejecuta JavaScript en el contexto de la pestaña autenticada mediante `chrome.scripting.executeScript` y `world: "MAIN"`. Las tareas incluyen catálogo, navegación, bloqueo/restauración, preparación, consulta, descarga y evidencia. Deben mantenerse abiertas la pestaña SITFA original y la de conexión durante el lote.

Pantallas documentadas: seguimiento, litigantes, calendario_informes, calendario_medidas y carga. La identificación se basa en formularios, campos y botones; los endpoints incluyen `InformesDAction.do` y `MaoDAction.do`. Estos nombres no bastan para reconstruir sus contratos completos.

`catalog` lee directamente los controles institucionales: tribunales, modalidades, informes, medidas, estados, meses, años y pestañas. En Seguimiento, `prepare` configura tribunal, pestaña —Espera, Cumplimiento, Informes o Egresados—, modalidad, tipo/filtro y rango de fechas cuando corresponde.

Regla crítica de V4: establecer filtros → llamar a `Envio()` → comprobar rechazo → reaplicar los valores alterados y limpiar campos no usados → construir `FormData`. Los inputs de identificación vacíos permanecen como cadena vacía; no se rellenan con cero. Los pares finales del POST, validados según campos permitidos e `irAccion`, son la fuente de verdad de la consulta.

La extensión ejecuta `fetch` con la sesión del navegador, sin caché, timeout y límite de tamaño. Valida respuestas y enlaces Excel restringidos al host institucional y rutas `/sitfa/reportes/<archivo>.xls`; transfiere los bytes a la aplicación local, que conserva el archivo institucional en la carpeta del lote.

Cuando hay cero resultados, genera PDF de evidencia: repite el POST final en un iframe, restaura visualmente los filtros desde esos mismos pares, serializa el DOM con valores vivos, calcula SHA-256, verifica que no cambió y captura la pestaña para construir el PDF. El hash controla consistencia del documento de evidencia durante ese flujo; por sí solo no acredita autenticidad institucional ni valida una consulta futura.

La aplicación organiza lotes por fecha/hora, muestra consultas/registros/archivos y permite pausar, continuar y cancelar conservando productos ya guardados. `lock` conserva el estado original y `unlock` intenta restaurarlo. No consta en el documento una garantía de reanudación después de cerrar o reiniciar la aplicación; esa capacidad sigue siendo una ampliación propuesta.

La combinación múltiple documentada incluye tribunales, pestañas, una modalidad, tipos/filtros y período. Seleccionar varias modalidades y cruces arbitrarios en un único plan será una ampliación del coordinador; no se presume que ya exista en la versión 2.6.0.

## 3. Diferencias y correcciones necesarias

| Hallazgo | Consecuencia actual | Decisión de diseño |
|---|---|---|
| Cobertura Meta 4 fija de 68 combinaciones | No sirve para cualquier selección de trabajo | Cobertura calculada desde un plan explícito; perfil Meta 4 completo conserva las 68 |
| PDF/imágenes clasificados solo por nombre/ruta | Un archivo no leído puede representar cero filas y terminar en SIN REGISTROS | Separar evidencia recibida de cero registros confirmado |
| Un candidato por combinación en Meta 4 | Puede descartar archivos que contienen fragmentos complementarios | Distinguir versiones del mismo reporte de particiones que deben sumarse |
| Consolidador copia todos los archivos | Puede duplicar registros si hay varias versiones de un reporte | Inventario y resolución previa de versiones |
| Coincidencia por persona vs fila idéntica | Un mismo NNA en varias medidas no es un duplicado eliminable | Señalar tres clases de coincidencia; conservar por defecto |
| Distintas reglas ante errores | Un flujo se detiene y otro produce resultados potencialmente incompletos | Validación común, vista previa y estado explícito de borrador/verificado |
| Ambas herramientas eligen una hoja por archivo | Un consolidado con dos hojas no debe tratarse como un reporte RUS común | Importador específico que lea ambas hojas del consolidado |
| Consolidación genérica descarta columnas ajenas a su esquema | Se pierde FECHA OBS. y cualquier extensión | Conservar columnas adicionales; mapear las necesarias sin borrar información |
| Registros.xlsx tiene ruta fija de salida | Una reejecución puede reemplazar las fechas editadas | Carpeta de ejecución y salidas versionadas; el consolidado importado es de solo lectura |
| Catálogo y reglas por prefijos incrustados en código | Limita otros tribunales/modalidades y puede clasificar prefijos ambiguos | Catálogo extensible, códigos RUS cuando existan y tabla de equivalencias revisable |
| Fechas inválidas pueden permanecer como texto | Se arrastra un dato no validado al producto | Validación estricta de fechas, sin borrar el valor original |

## 4. Diseño de la herramienta

Nombre de trabajo: **RUS Integral**. Aplicación local para Windows, con operaciones largas en segundo plano, Excel como intercambio y resultados agrupados por ejecución. Esta elección responde a los BAT, carpetas y edición Excel actuales; es una propuesta técnica, no un requisito observado de CSMP Integral.

### Cuatro rutas de trabajo

| Ruta | Entrada | Productos |
|---|---|---|
| Descargar | Ítem RUS y plan de combinaciones | Archivos originales, manifiesto y resumen de descargas |
| Consolidar | Descargas nuevas o carpeta existente | Consolidado Espera/Cumplimiento y controles separados |
| Meta 4 desde descargas | Descarga RUS o carpeta, perfil Meta 4 y período | Consolidado editable, controles y Meta 4 |
| Meta 4 desde mi consolidado | Excel propio con Espera/Cumplimiento y FECHA OBS. | Meta 4 con esas fechas y reporte de validación |

Se puede interrumpir el flujo después del consolidado, editarlo en Excel y retomarlo días después. No se exige repetir la descarga. Una descarga de otro ítem, como órdenes de búsqueda, sigue disponible aunque su esquema no permita dividirla en Espera/Cumplimiento: debe tener un adaptador propio antes de ofrecer esa consolidación. No se deduce el estado de cualquier tabla solo por su nombre.

### Selección múltiple y combinaciones

La unidad de trabajo es `(pantalla/ítem, tribunal, modalidad, pestaña, filtros del reporte, período/corte)`. Para consolidar Espera/Cumplimiento, la pestaña determina el estado. Los catálogos se leerán mediante la operación `catalog` de la extensión SITFA y admitirán equivalencias locales para importar Excel.

Dos modos complementarios:

- **Cruzar selección:** varios tribunales × varias modalidades × estados marcados. La interfaz muestra el número de tareas antes de ejecutar.
- **Combinaciones específicas:** tabla editable donde se agregan o retiran cruces individuales. Ejemplo: Concepción/FAE/Espera y Tomé/Residencia/Cumplimiento, sin descargar Concepción/Residencia ni Tomé/FAE.

El catálogo distingue lo permitido por cada ítem y tribunal. Una combinación no disponible se marca No aplica con causa, en lugar de fallar o contarse como faltante. Guardar perfiles permite repetir una selección mensual. Los filtros se mostrarán según el formulario detectado y sus opciones vigentes. Cada tarea tendrá su selección completa, incluidas varias modalidades a lo largo del lote.

### Descarga

Se conservará el mecanismo documentado: aplicación Windows + extensión Chrome + puente local con código temporal. El usuario inicia sesión en SITFA desde Chrome; la aplicación vincula esa sesión sin recibir credenciales. La operación `catalog` alimenta las opciones y el coordinador envía tareas individuales a la extensión. Se mantienen abiertas las pestañas SITFA y de conexión. El anexo detalla el contrato propuesto y las reglas V4 que deben preservarse.

Cada tarea registra pendiente, en curso, descargada, sin registros confirmado, fallida o cancelada. Verificar que el archivo terminó de descargarse, puede abrirse y corresponde al plan; no basta que aparezca un nombre en la carpeta. Ante sesión vencida, pausar y permitir reanudar. Reintentar fallos transitorios con límite y descargar únicamente los pendientes al retomar.

Conservar originales por ejecución y un manifiesto con filtros, fecha de extracción, archivo, huella y resultado. No guardar contraseñas en Excel, reportes o configuración en texto plano.

### Consolidación general

Leer todas las hojas seleccionadas y clasificar por fila cuando el archivo mezcle tribunales/modalidades. Detectar el esquema por contenido. En contradicciones de título/columnas o encabezados ambiguos, detener la clasificación de ese archivo y pedir resolución en la interfaz.

Preservar todos los tribunales y modalidades incluidos en el plan; las no reconocidas se muestran para mapear o conservar como Sin clasificar. La restricción FAE/Residencia corresponde únicamente al perfil Meta 4.

Producto principal: `Consolidado_<perfil>_<corte>_<version>.xlsx` con exactamente dos hojas de datos, **Espera** y **Cumplimiento**. Mantener columnas base compatibles, `FECHA OBS.`, identificación estable del registro y origen. Permitir exportación de compatibilidad si un consumidor exige un esquema estricto.

Producto de control separado: `Control_<ejecucion>.xlsx` con Inventario, Cobertura, Coincidencias, Incidencias y Conciliación. Así el consolidado queda dedicado a trabajar los registros. No eliminar filas por RIT: una causa puede tener varios NNA o varias medidas.

Tres categorías de coincidencia:

1. Archivo idéntico por huella: evitar doble lectura y registrar la omisión.
2. Fila idéntica en los datos de negocio: mostrar origen y permitir decidir; no borrar sin resolución.
3. Misma persona por RUT/nombre: señal informativa de múltiples apariciones, no prueba de duplicación.

Para varios reportes de una combinación, el usuario define si son versiones alternativas o particiones. Preferir la emisión/corte verificados al timestamp del archivo. Si no se puede distinguir, presentar el conflicto antes de certificar el resultado.

### Meta 4 desde descargas

El perfil completo conserva 17 tribunales, FAE/Residencia y Espera/Cumplimiento: 68 combinaciones conforme al código recibido. Un subconjunto produce una **vista parcial**, claramente identificada; no presenta tribunales excluidos como SIN REGISTROS.

Recorrido: definir período → validar descargas → resolver cobertura → generar consolidado → completar FECHA OBS. si se desea → previsualizar → generar Meta 4. Se permite generar borrador sin fechas para mantener el uso actual. El estado verificado exige resolver las incidencias obligatorias del perfil; una fecha vacía se comunica como pendiente y no se inventa.

El motor conserva las cuatro hojas, ocho columnas, bloques por tribunal, totales y formato de fechas observados. La plantilla es versionada y configurable. Hasta contrastarla con la plantilla institucional, se identifica como formato del paquete v2.4. No se añaden columnas al formato institucional sin una opción expresa.

### Meta 4 desde mi consolidado

1. Elegir .xls/.xlsx/.xlsm. Abrir solo para lectura.
2. Detectar Espera y Cumplimiento sin importar mayúsculas/tildes; si los nombres son distintos, ofrecer asignación de hojas. Leer ambas aunque una tenga cero filas.
3. Proponer mapeo de encabezados y permitir corregirlo. Aceptar `fecha obs`, `FECHA OBS.` y `Fecha de observación` con equivalencias explícitas; no confundirla con fecha de emisión.
4. Validar campos requeridos: RIT, TRIBUNAL, TIPO, SEXO, DERIVACIÓN, FEC. RESOLUCIÓN y, según estado, T ESPERA o FEC. INGRESO EFECTIVO. RUT/nombre no son necesarios para exportar las ocho columnas, aunque ayudan a identificar registros.
5. Validar FECHA OBS. como fecha de Excel, DD-MM-AAAA, DD/MM/AAAA o AAAA-MM-DD. Seriales Excel usan el calendario real del libro. Fechas imposibles y fórmulas sin valor guardado se presentan para corrección. Los vacíos se conservan y se cuentan; no se sustituyen por la fecha de hoy.
6. Mostrar cuántas filas irán a cada hoja Meta 4 y cuáles quedan fuera de su perfil. Todo registro no reconocido queda trazado con fila original y causa.
7. Generar Meta 4 transfiriendo FECHA OBS. fila por fila, sin recalcularla ni modificar el Excel de entrada.

Un consolidado propio no prueba que existan todas las descargas originales. Si no tiene un manifiesto válido, su cobertura de extracción se muestra **no acreditada**; puede producirse el informe desde las filas aportadas, con esa limitación en el control. Una hoja vacía solo prueba ausencia de filas en el archivo recibido.

### Conservación de observaciones al actualizar

Para nuevas descargas, usar identificador del registro/medida RUS si existe. En su ausencia, proponer una clave compuesta por tribunal, RIT, persona, derivación, fecha de resolución e ingreso; su unicidad debe comprobarse, no asumirse. Estado y emisión se guardan como atributos de la versión para poder reconocer un paso de Espera a Cumplimiento.

No copiar fechas a todas las filas del mismo RIT. Si hay más de una correspondencia posible, mostrar conflicto. En la primera versión basta garantizar importación directa y exportación fiel; la reconciliación automática entre meses es una ampliación posterior.

## 5. Pantallas y propuesta visual

La maqueta adjunta permite explorar selección de cruces, perfil Meta 4 y traslado de fechas ficticias. Es un documento de diseño ejecutable en el navegador, sin servicios externos.

Propuesta visual provisional: barra lateral azul oscuro, cabecera compacta, fondo claro, tarjetas de tareas, botones con verbo explícito, tablas amplias y estados con texto además de color. No se puede afirmar similitud con CSMP Integral sin una referencia. El diseño se separa en tokens para ajustar colores, tipografía, iconos, espaciado y disposición cuando se aporte esa referencia.

| Pantalla | Función principal |
|---|---|
| Inicio | Elegir una de las cuatro rutas y retomar una ejecución |
| Plan de trabajo | Ítem, tribunales, modalidades, estados, período y cruces específicos |
| Descargas | Progreso por combinación, errores y reanudar pendientes |
| Archivos y cobertura | Vista de faltantes, vacíos confirmados, conflictos y evidencia sin leer |
| Consolidado | Tabla filtrable, edición de FECHA OBS., incidencias y exportación |
| Importar mi consolidado | Mapeo de hojas/columnas y validación de fechas |
| Meta 4 | Vista previa de las cuatro hojas, fechas, conteos y generar |
| Historial y perfiles | Ejecuciones, productos, planes guardados y versiones de plantilla |

Acciones importantes muestran qué archivos se producirán y el estado de validación. El trabajo se guarda automáticamente a nivel de ejecución; cancelar conserva tareas ya descargadas y no publica productos finales a medio escribir.

## 6. Arquitectura propuesta

Aplicación local Python con PySide6 como propuesta de interfaz, reutilizando pandas/openpyxl/xlrd. SQLite guarda planes, ejecuciones e incidencias; los Excel y originales permanecen en carpetas visibles. La descarga conserva la arquitectura documentada del ejecutable, extensión Chrome MV3 y puente en 127.0.0.1. PySide6 sustituiría la interfaz Tkinter descrita; esa migración no debe alterar el protocolo de consulta o evidencia. La integración exacta requiere revisar el código disponible o comprobar un contrato soportado por el ejecutable existente.

```text
Interfaz local
    └── Coordinador de ejecución + trabajos en segundo plano
         ├── Puente local ↔ extensión Chrome ↔ pestaña SITFA autenticada
         │    └── consulta/descarga/evidencia → originales + manifiesto
         ├── Importador reportes → registros normalizados
         ├── Importador consolidado propio → registros + FECHA OBS.
         ├── Validador → cobertura + incidencias + decisiones
         ├── Exportador consolidado → Espera / Cumplimiento
         └── Perfil Meta 4 → cuatro hojas + controles
```

Modelo mínimo: Plan, Combinación, Ejecución, Descarga, ArchivoFuente, Registro, Observación, Incidencia, Decisión y Producto. Un registro retiene valores originales y normalizados, ubicación de origen y modalidad/estado; una observación retiene su fecha, procedencia y versión. Reproducir un producto exige el plan, las fuentes, el mapeo, las decisiones y la versión de plantilla.

Contratos orientativos:

```text
download(task, session) → DownloadResult
inspect_report(file, schema) → records + evidence + issues
import_consolidated(file, mapping) → records + observations + issues
validate(records, plan, profile) → coverage + reconciliation + issues
export_consolidated(records, output) → artifact
render_meta4(records, observations, template, emission) → artifact + counts
```

`render_meta4` no escanea carpetas ni escribe el consolidado. Esto permite usar indistintamente descargas normalizadas o el archivo propio.

Carpetas por ejecución:

```text
RUS Integral/
  perfiles/
  plantillas/
  ejecuciones/<fecha_hora_id>/
    originales/
    manifiesto.json
    trabajo/
    productos/Consolidado_...xlsx
    productos/META4_...xlsx
    productos/Control_...xlsx
    productos/Resumen_...txt
```

Guardar productos primero en archivo temporal y moverlos al destino cuando terminan. Nunca reemplazar silenciosamente un consolidado editado. Separar entradas/salidas evita releer productos de la otra herramienta, cuyos nombres de exclusión actualmente difieren.

## 7. Reglas de calidad y criterios de aceptación

- La cobertura corresponde únicamente a las combinaciones efectivamente seleccionadas. El perfil completo Meta 4 espera 68.
- Un PDF suelto no acredita cero casos automáticamente. La evidencia de cero generada por el descargador puede cubrir una combinación cuando su manifiesto vincula tarea, POST final, resultado y PDF, y se han validado esos datos. Faltante, no aplica y vacío confirmado son estados distintos.
- Un archivo con varios tribunales/modalidades distribuye sus filas correctamente y preserva la procedencia.
- Una consolidación general admite modalidades adicionales y tribunales fuera del catálogo Meta 4; el perfil Meta 4 informa exclusiones sin borrarlas del consolidado.
- Reportes con esquemas incompatibles requieren adaptador o mapeo; no se generan estados inventados para otros ítems RUS.
- Dos NNA con el mismo RIT se mantienen como dos registros. Coincidencias por persona no se eliminan automáticamente.
- El consolidado principal tiene Espera y Cumplimiento; controles separados explican cobertura, exclusiones y conflictos.
- Para cada registro importado y aceptado, la fecha de observación de salida coincide con la fecha de entrada normalizada. Vacíos continúan vacíos; inválidos requieren resolución.
- Importar un consolidado nunca lo sobrescribe. Se puede generar Meta 4 sin conexión a RUS y sin carpeta de descargas.
- Conciliación de importación: filas leídas = incluidas + fuera del plan/perfil + rechazadas/pendientes. Cualquier fila omitida tiene causa; la conciliación distingue omisiones documentadas de fallos de escritura.
- Conciliación Meta 4: total de filas en sus cuatro hojas = registros aceptados para ese producto. Reabrir la salida comprueba también contenido, identificadores internos y fechas, no solo totales.
- SIN REGISTROS se usa solo en bloques cuyo vacío esté acreditado. Para fuentes propias sin cobertura, permitir una declaración explícita del usuario y registrarla; de otro modo el bloque queda pendiente en el borrador.
- No se habilita estado verificado con conflictos obligatorios sin resolver. Exportar borrador sigue disponible y deja el estado visible en el nombre/resumen.
- La descarga se reanuda tras error o cierre sin repetir tareas verificadas del mismo plan/corte.

## 8. Plan de implementación

1. **Completar verificación:** usar la arquitectura documental del descargador como base; revisar los módulos fuente disponibles o comprobar interfaces soportadas de la distribución local, pantallas SITFA, referencia visual CSMP Integral y plantilla institucional. Obtener muestras anonimizadas y determinar códigos estables y significado operativo de FECHA OBS. No es necesario volver a aportar todo el ejecutable para continuar el diseño.
2. **Extraer motor común:** importadores separados, modelo normalizado, catálogos y políticas comunes de incidencias. Mantener pruebas útiles del motor Meta 4.
3. **Primera entrega funcional:** carpeta → consolidado flexible; consolidado propio → Meta 4 con FECHA OBS.; vista previa y productos versionados. Estas rutas no dependen de RUS.
4. **Integrar descarga:** selección de ítem/cruces, sesión, validación de archivos, manifiesto y reanudación. Validar en el entorno real de RUS.
5. **Integrar interfaz:** adaptar tokens y navegación a CSMP Integral, empaquetar para Windows y comprobar en un equipo de uso habitual.
6. **Pilotear:** comparar filas, fechas y formatos con productos manuales del mismo corte. Aceptar por exactitud, no solo por funcionamiento visual.

El mecanismo de acceso ya está documentado. Una estimación de implementación requiere comprobar la reutilización de los componentes existentes y precisar cuántos esquemas de ítems se deben soportar.

## 9. Funcionalidades afines y cambios concretos en los productos

Son ampliaciones opcionales. Se conservan las ocho columnas del formato Meta 4 salvo solicitud expresa de cambiarlo.

| Funcionalidad | Utilidad | Cambio concreto en el producto | Prioridad |
|---|---|---|---|
| Validación de FECHA OBS. antes de exportar | Evitar completar fechas recién en el informe | Meta 4 incluye las fechas ya revisadas; Control informa vacíos e inválidos | Base |
| Descargar solo faltantes/fallidos | Completar un corte sin repetir todo | Manifiesto distingue tareas nuevas y reutilizadas; cobertura muestra qué se completó | Alta |
| Comparador entre cortes | Detectar ingresos, egresos y cambios de estado | Nuevo Comparativo.xlsx con altas, bajas y cambios; no cambia las cuatro hojas institucionales | Alta |
| Reutilizar observaciones por registro | Evitar recapturar fechas al actualizar | Consolidado incorpora fechas conciliadas; Control muestra conflictos; Meta 4 las recibe tras validación | Alta |
| Cola de revisión priorizada | Organizar revisiones por días de espera o permanencia | Nuevo Pendientes.xlsx con motivo, prioridad y datos del caso; umbrales configurables y confirmados | Media |
| Alertas de actualización de fichas | Usar las fechas de fichas y FEC. OÍDO ya disponibles | Nuevo Control_actualizaciones.xlsx con campos faltantes o antigüedad según umbral | Media |
| Seguimiento de coincidencias entre estados | Revisar apariciones simultáneas de una persona | Control incluye coincidencias cruzadas con RIT, tribunal, derivación y origen; no elimina filas | Alta |
| Asignar responsable y estado de revisión | Coordinar la revisión del consolidado | Columnas opcionales RESPONSABLE, ESTADO REVISIÓN y FECHA REVISIÓN en el consolidado; Meta 4 sigue igual | Media |
| Registro de cambios y cierre de versión | Saber qué archivo se entregó y de dónde salió | Resumen de ejecución con versión, huellas, conteos, incidencias y decisiones | Alta |
| Perfiles guardados | Repetir combinaciones habituales | Productos con nombre de perfil y selección exacta en el control | Alta |
| Exportación PDF para revisión | Compartir una vista fija del resultado | PDF adicional con saltos por tribunal, encabezados y totales; requiere validar paginación | Media |
| Extracción de PDF/imágenes con revisión | Recuperar datos cuando no hay Excel | Filas propuestas con origen y nivel de validación; nunca declarar cero automáticamente | Posterior |
| Adaptadores para otros productos RUS | Reutilizar descarga e inventario | Productos específicos de órdenes de búsqueda/revisiones, con esquemas propios por definir | Posterior |

La mejora con mayor impacto inmediato es la ruta **consolidado propio → Meta 4**: elimina la recaptura de FECHA OBS. en el documento final y permite revisar la información antes de generar el producto.

## 10. Referencias del código estudiado

Rutas relativas al contenido de los ZIP originales:

| Evidencia | Archivo y funciones |
|---|---|
| Catálogo, columnas y formato de cuatro hojas | META4_Automatizado_v2_4/META4_Automatizado/scripts/generate_meta4.py: TRIBUNALS, SOURCE_HEADERS, SHEET_DEFS |
| Lectura, clasificación y candidatos | Mismo archivo: inspect_excel, inspect_zero_record_file, select_candidates, consolidated_rows |
| FECHA OBS. y regeneración | Mismo archivo: find_col, read_registros, rebuild, main |
| Escritura y verificación | Mismo archivo: write_registros_workbook, verify_meta_output, verify_registros_output |
| Consolidación genérica y abortar por error | Consolidador_Excel_Espera_Cumplimiento/consolidar_excels.py: extract_file, detect_status, run |
| Coincidencias por persona | Mismo archivo: find_duplicate_records |
| Selección visual de carpeta | Consolidador_Excel_Espera_Cumplimiento/selector_carpeta.py |
| Comprobaciones existentes | META4_Automatizado_v2_4/META4_Automatizado/tests/test_generate_meta4.py |

La referencia del descargador se entrega como `Fuente_Descargador_SITFA_2_6_0_V4.md`; sus secciones 4–11 fundamentan conexión/consulta/descarga y las secciones 14–16 fundamentan evidencia y restauración. El detalle de integración figura en `Anexo_integracion_Descargador_SITFA.md`. Los resultados de ejecución de comprobaciones se entregan en `Validacion_del_estudio.md`. Los paquetes originales y sus motores no se han modificado.
