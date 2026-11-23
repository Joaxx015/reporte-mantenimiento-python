# Tablero de mantenimiento en Power BI Desktop

## Conectar el Excel

1. Generá el reporte siguiendo el [README](../README.md).
2. En Power BI Desktop, elegí **Obtener datos → Libro de Excel** y seleccioná `salida/reporte_mantenimiento.xlsx`. Usá ese archivo para que la demo pueda actualizarse con la siguiente ejecución.
3. En el navegador, seleccioná las hojas **Ordenes** y **Errores**, y elegí **Transformar datos**. Si los encabezados aparecen como datos, usá **Usar la primera fila como encabezado**.
4. En **Ordenes**, asigná tipo fecha a `fecha_programada` y `fecha_corte`, número entero a `dias_atraso` y texto a las demás columnas. En **Errores**, asigná número entero a `numero_registro` y texto a las demás columnas.
5. Aplicá los cambios y mantené ambas tablas sin relación. Los filtros de equipo aplican a las órdenes válidas y la tarjeta de errores muestra todos los rechazos del archivo.

[Microsoft Learn · Conector de Excel en Power Query](https://learn.microsoft.com/es-es/power-query/connectors/excel)

## Crear indicadores

Creá cada medida por separado. Si Power BI usa punto y coma como separador DAX, reemplazá las comas de las funciones.

```dax
Ordenes validas = COALESCE(COUNTROWS(Ordenes), 0)
```

```dax
Ordenes vencidas =
COALESCE(
    CALCULATE(COUNTROWS(Ordenes), Ordenes[estado_calculado] = "Vencida"),
    0
)
```

```dax
Registros rechazados = COALESCE(COUNTROWS(Errores), 0)
```

| Visual | Campos |
| --- | --- |
| Tres tarjetas | Una medida por tarjeta. |
| Gráfico de barras | `Ordenes[equipo]` como categoría y `Ordenes vencidas` como valor. |
| Tabla de detalle | Orden, equipo, tarea, fecha programada y días de atraso. Filtrar por estado `Vencida` y ordenar por atraso descendente. Seleccionar **No resumir** para `dias_atraso`. |

## Probar y dejar evidencia

1. Con el ejemplo original, esperá **4 válidas, 2 vencidas y 1 rechazada**.
2. Cambiá el estado de OT-001 a `Completada`, regenerá el Excel en la misma ruta y presioná **Actualizar** en Power BI. El resultado esperado es **4 válidas, 1 vencida y 1 rechazada**.
3. Volvé el CSV a su estado original. Guardá el tablero en `powerbi/mantenimiento.pbix` y agregá una captura a `docs/`.
4. Si movés el proyecto a otra PC, actualizá el origen de datos del PBIX para que apunte al Excel generado allí.
