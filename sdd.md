# SDD — Automatización de reportes de mantenimiento

**Autor del proyecto:** Joaquín Silva · Ingeniería en Informática, UNDAV  
**Estado:** script y Excel de ejemplo implementados. Tablero pendiente de armado en Power BI Desktop · Versión 1.0

## 1. Problema y objetivo

En un caso ficticio, el área de mantenimiento recibe un CSV con órdenes de trabajo y revisa manualmente qué tareas están vencidas. El proyecto automatiza esa revisión con Python, genera un Excel y permite consultar los resultados en Power BI Desktop.

La pregunta que responde es **qué equipos tienen tareas de mantenimiento vencidas y cuáles requieren atención primero**. El demo usa datos ficticios que se pueden publicar en el repositorio.

## 2. Alcance

| Incluido en la primera versión | Extensiones posibles |
| --- | --- |
| Leer un CSV local con órdenes de mantenimiento | Importar varias fuentes o integrar un sistema de mantenimiento |
| Validar los registros y calcular el atraso a una fecha de corte | Incorporar responsables, criticidad o historial de cierres |
| Generar un Excel con órdenes válidas y registros rechazados | Incorporar inventario de repuestos |
| Actualizar manualmente un tablero de Power BI Desktop | Programar actualizaciones o publicar el tablero en la nube |

## 3. Diseño del proceso

```text
CSV de órdenes → Script Python → Excel con dos hojas → Power BI Desktop
```

| Componente | Responsabilidad |
| --- | --- |
| `procesar_ordenes.py` | Lee los argumentos y el CSV, valida las filas, aplica las reglas y exporta el Excel. |
| `reporte_mantenimiento.xlsx` | Presenta el resultado y funciona como origen de datos del tablero. |
| `mantenimiento.pbix` | Presenta indicadores y el detalle de órdenes vencidas por equipo. |

Usar Python 3.11 o posterior. La biblioteca estándar cubre los argumentos, el CSV, las fechas y las rutas. Usar `pandas` y `openpyxl` para escribir el Excel. Una vez probado el flujo, registrar en `requirements.txt` las versiones con las que funcionó.

### Estructura objetivo del repositorio

La versión inicial incluye también una guía de Power BI y un Excel generado en `docs/`.

```text
reporte-mantenimiento-python/
├── README.md
├── sdd.md
├── procesar_ordenes.py
├── requirements.txt
├── .gitignore
├── datos/
│   └── ordenes_mantenimiento.csv
├── powerbi/
│   └── mantenimiento.pbix
└── docs/
    └── tablero.png
```

El script genera `salida/reporte_mantenimiento.xlsx`. Agregar `salida/`, `.venv/` y `__pycache__/` a `.gitignore`. Usar rutas relativas a la carpeta del proyecto en los ejemplos del README.

### Comando previsto

Desde la raíz del repositorio, con Python y las dependencias instaladas:

```powershell
python -m pip install -r requirements.txt
python procesar_ordenes.py --entrada datos/ordenes_mantenimiento.csv --salida salida/reporte_mantenimiento.xlsx --fecha-corte 2026-11-22
```

La fecha de corte es obligatoria y se escribe como `AAAA-MM-DD`. Así se puede reproducir la demo o revisar el reporte a una fecha anterior.

## 4. Contrato del CSV

Archivo UTF-8, con comas como separador y una fila de encabezados. Se aceptan espacios al comienzo y al final de los valores y se limpian antes de validar. Todos los campos son obligatorios.

| Columna | Tipo | Ejemplo o valor admitido |
| --- | --- | --- |
| `orden_id` | Texto | `OT-001`. Identifica una orden. |
| `equipo` | Texto | `Esmaltadora` |
| `tarea` | Texto | `Revisar boquillas` |
| `fecha_programada` | Fecha `DD/MM/AAAA` | `18/11/2026` |
| `estado_original` | Texto | `Pendiente` o `Completada`. Se aceptan distintas combinaciones de mayúsculas y minúsculas. |

### Datos ficticios del demo

```csv
orden_id,equipo,tarea,fecha_programada,estado_original
OT-001,Esmaltadora,Revisar boquillas,18/11/2026,Pendiente
OT-002,Esmaltadora,Limpiar filtros,20/11/2026,Pendiente
OT-003,Horno,Revisar sensores,25/11/2026,Pendiente
OT-004,Cinta transportadora,Ajustar tensión,19/11/2026,Completada
OT-005,Horno,Inspección general,31/11/2026,Pendiente
```

## 5. Reglas de negocio y validación

1. Si la orden está completada, el estado calculado será **Completada** y el atraso será 0 días. El CSV carece de fecha real de cierre, por lo que el indicador mide únicamente el atraso de las órdenes abiertas.
2. Si la orden está pendiente y la fecha programada es anterior al corte, será **Vencida**. Los días de atraso son la diferencia en días corridos entre ambas fechas.
3. Si la orden está pendiente y la fecha programada es igual o posterior al corte, será **Pendiente** y el atraso será 0 días.
4. Si la fila tiene campos faltantes, una fecha inválida, un estado desconocido o una cantidad de valores distinta de la esperada, se registra en **Errores**. Se guarda el primer error detectado por fila y se continúa con la siguiente.
5. Si un identificador se repite, se conserva la primera fila válida y las filas válidas posteriores con ese identificador se registran como duplicadas.

### Flujo del script

1. Validar los argumentos, abrir el CSV y verificar los encabezados requeridos. Si el archivo está vacío y carece de encabezados, hay columnas requeridas ausentes o el CSV es ilegible, mostrar el motivo, terminar con error y conservar cualquier Excel anterior.
2. Recorrer las filas y construir dos listas de registros: válidos y rechazados. Implementar la clasificación en una función que reciba estado, fecha programada y corte, y devuelva estado calculado y días de atraso.
3. Exportar el Excel con sus dos hojas y encabezados fijos, incluso si alguna queda sin registros. Un CSV con encabezados y sin filas genera ambas hojas vacías.
4. Escribir primero en un archivo temporal de la carpeta de salida. Reemplazar el Excel anterior después de terminar de escribir ambas hojas. Si el archivo está bloqueado, indicar que se cierre en Excel, conservar el resultado previo y limpiar el temporal.
5. Mostrar la fecha de corte, el total de órdenes válidas, vencidas y rechazadas, y la ruta del archivo generado. Las filas rechazadas cuentan como parte de una ejecución exitosa si se pudo exportar el reporte.

## 6. Contrato del Excel

| Hoja | Columnas, en orden | Tipo y significado |
| --- | --- | --- |
| `Ordenes` | `orden_id`, `equipo`, `tarea`, `fecha_programada`, `estado_original`, `estado_calculado`, `dias_atraso`, `fecha_corte` | IDs y etiquetas como texto, ambas fechas como celdas de fecha de Excel, atraso como entero. |
| `Errores` | `numero_registro`, `orden_id`, `campo`, `valor_recibido`, `motivo` | Número de registro entero contando los registros de datos desde 1. Los demás campos son texto. |

Cada hoja empieza con los encabezados en la primera fila. Mantener los nombres de columnas entre ejecuciones. Formatear las fechas como `DD/MM/AAAA` para su lectura en Excel.

### Resultado esperado con corte al 22/11/2026

| Orden | Equipo | Estado calculado | Días de atraso |
| --- | --- | --- | ---: |
| OT-001 | Esmaltadora | Vencida | 4 |
| OT-002 | Esmaltadora | Vencida | 2 |
| OT-003 | Horno | Pendiente | 0 |
| OT-004 | Cinta transportadora | Completada | 0 |

En `Errores` debe figurar el registro 5, orden `OT-005`, campo `fecha_programada`, valor `31/11/2026` y motivo `Fecha inválida`. El total esperado es **4 órdenes válidas, 2 vencidas y 1 registro rechazado**.

## 7. Power BI Desktop

1. Generar primero el Excel. En Power BI Desktop, conectarse al archivo mediante **Obtener datos → Libro de Excel**. Seleccionar `Ordenes` y `Errores`.
2. En Power Query, usar la primera fila como encabezados si corresponde. Confirmar tipos de fecha, texto y número entero. Nombrar las consultas `Ordenes` y `Errores`.
3. Mantener las dos tablas independientes. El filtro por equipo aplica a las órdenes válidas. La tarjeta de rechazos muestra el total de registros rechazados en el CSV de la ejecución actual, independientemente del equipo seleccionado.
4. Para repetir el demo en la misma PC, regenerar el archivo en la misma ruta y seleccionar **Actualizar**. Si otra persona abre el PBIX, debe apuntar el origen al Excel generado en su equipo.

| Visual | Dato |
| --- | --- |
| Tarjeta **Órdenes válidas** | Cantidad de filas de `Ordenes`. Incluye vencidas, pendientes y completadas. |
| Tarjeta **Órdenes vencidas** | Filas de `Ordenes` con estado calculado igual a `Vencida`. |
| Tarjeta **Registros rechazados** | Cantidad de filas de `Errores`. |
| Barras **Vencidas por equipo** | Equipo en el eje y cantidad de órdenes vencidas como valor. |
| Tabla **Órdenes vencidas** | Orden, equipo, tarea, fecha programada y días de atraso. Orden descendente por atraso. |

Medidas sugeridas. Si la configuración regional lo requiere, usar punto y coma como separador de argumentos en DAX.

```dax
Ordenes validas = COALESCE(COUNTROWS(Ordenes), 0)

Ordenes vencidas =
COALESCE(
    CALCULATE(COUNTROWS(Ordenes), Ordenes[estado_calculado] = "Vencida"),
    0
)

Registros rechazados = COALESCE(COUNTROWS(Errores), 0)
```

[Microsoft Learn — Conector de libros de Excel en Power Query](https://learn.microsoft.com/en-us/power-query/connectors/excel)

## 8. Criterios de aceptación y pruebas

| Caso | Resultado esperado |
| --- | --- |
| Ejecutar con las cinco filas de ejemplo y corte al 22/11/2026 | Excel con 4 órdenes y 1 error. Power BI muestra 2 vencidas, ambas de Esmaltadora. |
| Fecha programada igual al corte y estado pendiente | Estado calculado `Pendiente`, con 0 días de atraso. |
| Cambiar OT-001 a `Completada`, ejecutar de nuevo y actualizar el tablero | 4 órdenes válidas, 1 vencida y 1 registro rechazado. |
| Agregar un ID repetido después de su primera fila válida | La orden cuenta una sola vez y el duplicado aparece en `Errores`. |
| Corregir la fecha de OT-005 a 30/11/2026 | 5 órdenes válidas, 2 vencidas y 0 rechazadas. `Errores` conserva sus encabezados. |
| Quitar una columna obligatoria del CSV | La ejecución indica qué columna falta y conserva el Excel anterior. |
| Dejar el archivo de salida abierto de modo que su reemplazo resulte bloqueado | La ejecución explica el problema y conserva el Excel anterior. |

El proyecto queda listo para mostrar cuando Joaquín puede ejecutar el script desde el README, obtener los resultados de la muestra y actualizar el tablero después de cambiar una orden.

## 9. Implementación en dos sesiones

| Sesión | Trabajo de Joaquín | Hito |
| --- | --- | --- |
| **1. Leer y clasificar** | Crear el CSV, recorrer las filas, convertir fechas, escribir la función de clasificación y validar los registros. | Los conteos correctos aparecen en consola y la fecha inválida queda identificada. |
| **2. Exportar y visualizar** | Escribir las hojas del Excel, conectar Power BI y probar la actualización. | El indicador de vencidas baja de 2 a 1 cuando se completa una orden. |

## 10. Presentación en GitHub

**Nombre sugerido:** `reporte-mantenimiento-python`

**Descripción del repositorio:** Automatización de órdenes de mantenimiento con Python, Excel y Power BI. Datos ficticios para una demo reproducible.

Inicio de README para usar una vez que funcione el proyecto:

> # Reporte de mantenimiento con Python, Excel y Power BI
>
> Proyecto personal de Joaquín Silva. Procesa órdenes de mantenimiento desde un CSV, identifica tareas vencidas y genera un Excel que alimenta un tablero de Power BI.
>
> **Caso de ejemplo.** Al 22/11/2026, el reporte identifica 2 órdenes vencidas en Esmaltadora entre 4 órdenes válidas y registra 1 fila con fecha inválida.

Debajo del inicio, agregar en este orden:

1. Una captura real del tablero con los datos de ejemplo.
2. Dos comandos para instalar las dependencias y ejecutar el script.
3. Resultado esperado, ubicación del Excel y pasos para actualizar Power BI.
4. Enlace a este SDD y a un video breve, si está disponible.

**Guion para la demo.** Mostrar que Esmaltadora tiene dos órdenes vencidas. Cambiar OT-001 a completada, ejecutar Python y actualizar Power BI. Mostrar que queda una vencida y explicar por qué el registro con fecha inválida sigue en la hoja de errores.
