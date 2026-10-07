# SITFA Descargador 2.6.0 — Funcionamiento interno y flujo completo

Documento técnico de referencia para programadores y sistemas de IA que necesiten comprender o modificar el producto.

Versión de referencia: **2.6.0 corregida V4**.

## 1. Qué es el Descargador SITFA

Es una aplicación Windows portable que automatiza consultas y descargas desde determinadas pantallas de SITFA utilizando la sesión autenticada que el usuario ya mantiene abierta en Google Chrome.

No automatiza el login ni necesita recibir la contraseña del usuario.

Su arquitectura se divide en dos componentes principales:

1. **Aplicación Windows** (`SITFA_Descargador.exe`).
2. **Extensión Chrome** (`_internal/extension`).

Ambos se comunican localmente mediante `127.0.0.1` y un código temporal de conexión.

## 2. Componentes necesarios

### 2.1 Aplicación Windows

Archivo principal:

`SITFA_Descargador.exe`

Distribución:

- ejecutable x64;
- PyInstaller;
- Python 3.12 embebido;
- dependencias dentro de `_internal`.

El ejecutable no debe copiarse aisladamente.

### 2.2 Carpeta `_internal`

Contiene:

- `python312.dll`;
- librerías Python y DLL necesarias;
- Tcl/Tk para la interfaz;
- extensión Chrome;
- otros recursos requeridos por PyInstaller.

### 2.3 Extensión Chrome

Ruta:

`_internal/extension`

Archivos relevantes:

- `manifest.json`;
- `background.js`;
- `conexion.html`;
- `conexion.css`;
- `conexion.js`;
- `pagina.js`.

Manifest:

- Manifest V3;
- nombre visible: `SITFA · Descargador local`;
- versión: `2.6.0`;
- permisos: `activeTab`, `scripting`;
- acceso a `https://familia.pjud.cl/*`;
- acceso al puente local `http://127.0.0.1/*`.

### 2.4 Navegador y sesión

Requiere:

- Google Chrome o entorno Chromium compatible con Manifest V3;
- extensión cargada como extensión descomprimida;
- sesión SITFA ya iniciada;
- pestaña SITFA original abierta durante la ejecución;
- pestaña de conexión de la extensión abierta durante el lote.

## 3. Modelo de seguridad y separación de responsabilidades

La aplicación local no necesita recibir directamente cookies de SITFA.

La extensión ejecuta funciones JavaScript en el contexto de la pestaña autenticada usando `chrome.scripting.executeScript` con `world: "MAIN"`.

Esto permite interactuar con formularios y funciones JavaScript de SITFA dentro del contexto ya autenticado del navegador.

El puente local utiliza un código temporal que cambia al reiniciar o desconectar el Descargador.

El objetivo del diseño es que:

- la autenticación permanezca en Chrome;
- el ejecutable coordine el trabajo;
- la extensión haga las operaciones que necesitan contexto de navegador.

## 4. Flujo de conexión

### Paso 1. Inicio del Descargador

El usuario ejecuta `SITFA_Descargador.exe`.

La aplicación levanta un servicio local en `127.0.0.1` y muestra un código de conexión temporal con formato equivalente a:

`PUERTO-CLAVE`

### Paso 2. Apertura de SITFA

El usuario inicia sesión normalmente en SITFA y abre una pantalla compatible, normalmente Seguimiento.

### Paso 3. Apertura de la extensión

Desde la pestaña SITFA, el usuario pulsa el icono de la extensión.

La extensión abre su página de conexión asociada a la pestaña original.

### Paso 4. Vinculación

El usuario copia el código desde la aplicación Windows y lo pega en la extensión.

`conexion.js` valida el formato y se conecta a:

`http://127.0.0.1:<puerto>`

El encabezado local incluye, entre otros:

- código temporal;
- ID de la extensión.

### Paso 5. Bucle de tareas

Una vez conectada, la extensión consulta periódicamente al Descargador por nuevas tareas (`/next`).

El Descargador puede solicitar operaciones como:

- catálogo;
- navegación;
- bloqueo de pantalla;
- preparación de consulta;
- ejecución de consulta;
- descarga;
- apertura/captura/cierre de evidencia.

La extensión devuelve el resultado por el puente local.

## 5. Detección de pantalla SITFA

`pagina.js` identifica el formulario disponible y clasifica la pantalla.

Entre las pantallas reconocidas están:

- `seguimiento`;
- `litigantes`;
- `calendario_informes`;
- `calendario_medidas`;
- `carga`.

La identificación se basa en:

- nombre del formulario;
- campos presentes;
- botones o elementos característicos.

Los endpoints principales identificados son del tipo:

- `InformesDAction.do`;
- `MaoDAction.do`.

## 6. Carga de catálogos y opciones

Cuando el usuario pulsa `Cargar opciones`, la extensión ejecuta la acción `catalog`.

La función lee opciones disponibles directamente desde los controles de SITFA, por ejemplo:

- tribunales;
- modalidades;
- tipos de informe;
- medidas;
- estados;
- meses;
- años;
- pestañas disponibles.

La aplicación Windows usa esos datos para construir la interfaz del lote.

Esto evita mantener catálogos rígidos independientes cuando SITFA ya expone las opciones actuales.

## 7. Preparación de una consulta

La operación central en la extensión es `prepare`.

Recibe una selección estructurada desde el Descargador y prepara el formulario real de SITFA.

### 7.1 En Seguimiento

Puede configurar, según corresponda:

- tribunal;
- pestaña (`Espera`, `Cumplimiento`, `Informes`, `Egresados`);
- modalidad;
- tipo de informe/filtro;
- rango de fechas;
- filtros adicionales conservados desde la sesión original.

### 7.2 Rango de fechas

Cuando se utiliza rango:

- se marca el control de filtrado;
- se habilitan `FEC_Inicio` y `FEC_Fin`;
- se establecen las fechas solicitadas.

### 7.3 Filtros no utilizados

Regla importante de V4:

- inputs de texto no utilizados deben quedar `""`;
- no se debe utilizar `0` como sustituto genérico de vacío;
- selects usan únicamente una opción neutral válida para ese control;
- radio/checkbox se marcan o desmarcan según corresponda.

Especial atención a:

- RIT;
- RUT;
- dígito verificador;
- otros inputs de identificación.

## 8. Interacción con JavaScript nativo de SITFA

SITFA posee funciones JavaScript propias que preparan o validan el formulario.

En Seguimiento, una función relevante es `Envio()`.

Este punto es crítico porque `Envio()` puede modificar el estado que la extensión había preparado.

### Regla V4

El flujo correcto es:

1. configurar valores deseados;
2. llamar a `Envio()`;
3. comprobar si SITFA rechazó la consulta;
4. **reaplicar los valores que SITFA haya modificado**;
5. limpiar nuevamente inputs que deben permanecer vacíos;
6. recién después construir el POST.

No invertir este orden.

## 9. Construcción del POST

Después de estabilizar el formulario se utiliza `FormData` sobre el formulario y el botón correspondiente.

El resultado se convierte a una lista de pares:

`[[nombre, valor], ...]`

Estos pares deben cumplir un conjunto de campos permitidos para cada tipo de pantalla.

Además se valida que:

- no existan nombres duplicados inesperados;
- `irAccion` corresponda a la acción prevista para esa pantalla;
- no se envíen campos ajenos al contrato reconocido.

### Fuente de verdad

En V4, esta lista de pares representa la **consulta efectivamente enviada**.

Toda evidencia posterior debe reconstruirse desde estos pares.

## 10. Ejecución de la consulta

La extensión ejecuta el POST utilizando `fetch` con:

- `credentials: "include"`;
- `cache: "no-store"`;
- timeout acotado;
- sesión del navegador ya autenticada.

La respuesta se lee como bytes con límite máximo de tamaño.

No se acepta silenciosamente una respuesta vacía o estructuralmente inesperada.

## 11. Descarga de Excel

Cuando SITFA devuelve un enlace a un reporte Excel, la extensión valida la URL antes de descargarla.

El patrón permitido está limitado a recursos del propio host SITFA, en rutas del tipo:

`/sitfa/reportes/<archivo>.xls`

No se permiten arbitrariamente hosts externos ni URLs con parámetros inesperados.

La descarga utiliza la misma sesión autenticada del navegador.

Los bytes resultantes son transferidos al Descargador local, que los guarda en la carpeta del lote.

## 12. Organización del lote

El usuario selecciona una carpeta base de salida.

La aplicación crea una subcarpeta identificada por fecha/hora para cada ejecución.

Según las opciones seleccionadas, puede realizar múltiples consultas combinando, por ejemplo:

- varios tribunales;
- Espera y Cumplimiento;
- una modalidad;
- filtros/tipos;
- período temporal.

La interfaz muestra durante el lote información como:

- página;
- cantidad de registros;
- archivo guardado.

## 13. Productos que entrega

### 13.1 Archivos Excel

Cuando SITFA tiene resultados, el producto principal es el archivo `.xls` generado por el propio sistema institucional.

El Descargador no debe reinterpretar el contenido del archivo para “inventar” resultados: conserva el archivo entregado por SITFA.

### 13.2 PDF de evidencia

Cuando una consulta devuelve cero registros y corresponde documentar ese resultado, se genera un PDF a partir de una captura visual de SITFA.

La finalidad es acreditar:

- qué pantalla fue consultada;
- qué filtros se utilizaron;
- que el resultado visible fue cero registros.

### 13.3 Tabla/ventana de resultados

La aplicación mantiene un resumen del lote con información sobre consultas, registros y archivos generados.

### 13.4 Carpeta completa del lote

El usuario obtiene una carpeta autocontenida con los productos de la ejecución.

## 14. Generación de evidencia PDF — arquitectura V4

Este es el punto más delicado de la versión actual.

### 14.1 Problema que resuelve

Una respuesta vacía de SITFA puede mostrar valores por defecto diferentes de los enviados.

Capturar esa página directamente genera una evidencia falsa o incompleta.

### 14.2 Flujo de evidencia

1. La extensión recibe el POST final usado para la consulta.
2. Crea un panel a pantalla completa.
3. Inserta un iframe destinado exclusivamente a evidencia.
4. Ejecuta en ese iframe un POST nativo al endpoint SITFA usando los mismos pares.
5. Espera la carga completa.
6. Obtiene el documento del iframe.
7. Ejecuta `restoreEvidenceSelection`.
8. La función vuelve a aplicar visualmente:
   - pestaña;
   - tribunal;
   - modalidad;
   - tipo;
   - checkbox de fecha;
   - fechas;
   - demás controles relevantes.
9. Si SITFA dejó un campo de fecha desactivado, la evidencia lo vuelve a habilitar para que el valor sea visible.
10. Se serializa el DOM incluyendo los valores vivos de inputs/selects/textareas.
11. Se calcula SHA-256 de la estructura reconstruida.
12. `evidence_show` comprueba que el documento no haya cambiado desde la validación.
13. Se oculta la cubierta del Descargador.
14. `conexion.js` activa la pestaña SITFA y utiliza `chrome.tabs.captureVisibleTab`.
15. La captura PNG se devuelve al ejecutable.
16. El ejecutable incorpora la captura al PDF correspondiente.
17. El iframe/panel de evidencia se cierra y la pantalla original permanece sin modificación permanente.

## 15. Por qué la evidencia no debe depender del DOM devuelto por SITFA

La respuesta HTML no es una fuente confiable del estado de la consulta porque SITFA puede:

- reiniciar selects;
- cambiar pestañas;
- reponer fechas;
- desmarcar filtros;
- mostrar valores predeterminados.

Por eso V4 adopta esta regla:

> POST efectivamente enviado = fuente de verdad para reconstruir la evidencia.

Esta regla es una decisión de diseño y debe mantenerse.

## 16. Bloqueo y restauración de la pantalla original

Durante un lote, la extensión puede ejecutar `lock`.

Antes de modificar el formulario conserva un snapshot de los campos y de la pestaña activa.

Al finalizar utiliza `unlock` para intentar restaurar:

- valores originales;
- estados checked;
- disabled;
- pestaña activa.

El objetivo es reducir el impacto visible de la automatización en la sesión interactiva del usuario.

## 17. Pausa y cancelación

La aplicación permite pausar/continuar y cancelar el lote.

Cancelar no debe eliminar productos ya verificados y guardados.

La ejecución puede terminar conservando el trabajo parcial disponible.

## 18. Restricciones y supuestos técnicos

- SITFA puede cambiar HTML, nombres de campos o funciones JavaScript sin aviso.
- La extensión depende de estructuras concretas del formulario.
- Si cambia `Envio()`, debe repetirse la prueba de estabilización posterior a esa función.
- Si cambia un endpoint, la lista de campos permitidos debe revisarse.
- La extensión está diseñada para `familia.pjud.cl`; no debe ampliarse host permissions sin necesidad explícita.
- El ejecutable original 2.6.0 no venía acompañado de todo su código fuente exacto.

## 19. Prueba mínima de regresión recomendada

Antes de distribuir una nueva versión comprobar, al menos:

1. Seguimiento → Espera → consulta con resultados.
2. Seguimiento → Cumplimiento → consulta con resultados.
3. Consulta con rango de fechas.
4. Consulta sin resultados → PDF con filtros correctos.
5. RIT y RUT vacíos siguen vacíos.
6. Fechas no se reemplazan por la fecha del día después de `Envio()`.
7. Tribunal/modalidad no se pierden después de `Envio()`.
8. Excel descargado abre correctamente.
9. Cancelar conserva archivos previos.
10. `unlock` devuelve la pantalla original a un estado razonablemente equivalente al inicial.
