"""Clasifica órdenes de mantenimiento y genera un Excel para Power BI."""

# Permite interpretar los argumentos recibidos desde la terminal.
import argparse
# Permite leer y procesar archivos con formato CSV.
import csv
# Permite mostrar errores en la salida estándar de errores.
import sys
# Permite representar fechas y convertir fechas entre diferentes formatos.
from datetime import date, datetime
# Permite trabajar con rutas de archivos de manera independiente del sistema operativo.
from pathlib import Path
# Permite crear un archivo temporal para exportar el Excel de forma segura.
from tempfile import NamedTemporaryFile

# Permite crear tablas de datos y archivos Excel.
import pandas as pd


# Define las columnas obligatorias que debe contener el archivo CSV.
COLUMNAS_CSV = ("orden_id", "equipo", "tarea", "fecha_programada", "estado_original")
# Define las columnas adicionales que tendrá la hoja de órdenes exportada.
COLUMNAS_ORDENES = (*COLUMNAS_CSV, "estado_calculado", "dias_atraso", "fecha_corte")
# Define las columnas que tendrá la hoja de errores exportada.
COLUMNAS_ERRORES = ("numero_registro", "orden_id", "campo", "valor_recibido", "motivo")
# Relaciona los estados aceptados sin distinguir mayúsculas y minúsculas con sus nombres normalizados.
ESTADOS_VALIDOS = {"pendiente": "Pendiente", "completada": "Completada"}


def clasificar_orden(estado, fecha_programada, fecha_corte):
    """Clasifica una orden como completada, vencida o pendiente según sus fechas."""

    # Devuelve el estado completada y cero días de atraso cuando la orden está finalizada.
    if estado == "Completada":
        return "Completada", 0

    # Calcula la cantidad de días transcurridos desde la fecha programada hasta la fecha de corte.
    atraso = (fecha_corte - fecha_programada).days
    # Determina que la orden está vencida cuando el atraso supera cero días.
    if atraso > 0:
        return "Vencida", atraso
    # Devuelve pendiente para las órdenes que todavía están dentro de su plazo.
    return "Pendiente", 0


def leer_fecha_programada(valor):
    """Convierte una fecha en texto al formato interno de fecha y valida su escritura."""

    # Interpreta el texto recibido como una fecha escrita con día, mes y año.
    fecha = datetime.strptime(valor, "%d/%m/%Y").date()
    # Comprueba que el texto tenga exactamente el formato esperado sin espacios ni variaciones.
    if valor != fecha.strftime("%d/%m/%Y"):
        raise ValueError("Formato de fecha inválido")
    # Devuelve la fecha convertida para que pueda utilizarse en cálculos.
    return fecha


def validar_orden(orden, ids_validos):
    """Valida los campos y el identificador de una orden antes de incorporarla al reporte."""

    # Revisa cada columna obligatoria para comprobar que contiene un valor.
    for campo in COLUMNAS_CSV:
        # Si una columna obligatoria está vacía, devuelve el detalle de ese error.
        if not orden[campo]:
            return campo, orden[campo], "Campo obligatorio"

    # Intenta interpretar la fecha programada usando la función específica para ese campo.
    try:
        fecha_programada = leer_fecha_programada(orden["fecha_programada"])
    # Convierte cualquier error de formato de fecha en un resultado de validación.
    except ValueError:
        return "fecha_programada", orden["fecha_programada"], "Fecha inválida"

    # Obtiene el estado escrito originalmente en el registro.
    estado_original = orden["estado_original"]
    # Busca una versión normalizada del estado sin distinguir mayúsculas y minúsculas.
    estado_normalizado = ESTADOS_VALIDOS.get(estado_original.casefold())
    # Rechaza el registro si el estado no coincide con ninguno de los valores admitidos.
    if estado_normalizado is None:
        return "estado_original", estado_original, "Estado desconocido"

    # Extrae el identificador de la orden para comprobar si ya fue procesada.
    orden_id = orden["orden_id"]
    # Devuelve un error cuando el identificador ya pertenece a otra orden válida.
    if orden_id in ids_validos:
        return "orden_id", orden_id, "Orden duplicada"

    # Reemplaza el texto de la fecha por el objeto de fecha validado.
    orden["fecha_programada"] = fecha_programada
    # Reemplaza el estado original por su forma normalizada.
    orden["estado_original"] = estado_normalizado
    # Indica que la orden no presentó errores de validación.
    return None


def procesar_csv(ruta, fecha_corte):
    """Lee el CSV, valida sus registros y devuelve las órdenes válidas y los errores encontrados."""

    # Inicializa la lista que almacenará las órdenes procesadas correctamente.
    ordenes = []
    # Inicializa la lista que almacenará los registros rechazados.
    errores = []
    # Inicializa el conjunto que permitirá detectar identificadores duplicados.
    ids_validos = set()

    # Abre el archivo CSV con codificación UTF-8 y sin transformar automáticamente las líneas.
    with ruta.open(encoding="utf-8-sig", newline="") as archivo:
        # Crea un lector CSV estricto que pueda detectar errores de formato.
        lector = csv.reader(archivo, strict=True)
        # Lee la primera fila como encabezados o una lista vacía si el archivo no tiene contenido.
        encabezados = next(lector, [])
        # Rechaza el archivo si no contiene una fila de encabezados.
        if not encabezados:
            raise ValueError("El CSV está vacío o no tiene encabezados.")
        # Rechaza el archivo si algún nombre de columna aparece más de una vez.
        if len(set(encabezados)) != len(encabezados):
            raise ValueError("El CSV tiene encabezados duplicados.")

        # Obtiene la lista de columnas obligatorias que no aparecen en el encabezado.
        faltantes = [campo for campo in COLUMNAS_CSV if campo not in encabezados]
        # Rechaza el archivo si falta al menos una columna obligatoria.
        if faltantes:
            raise ValueError(f"Faltan columnas obligatorias: {', '.join(faltantes)}.")

        # Guarda la posición de la columna que contiene el identificador de cada orden.
        indice_id = encabezados.index("orden_id")
        # Recorre las filas restantes comenzando la numeración desde el registro uno.
        for numero, fila in enumerate(lector, start=1):
            # Elimina espacios al inicio y al final de cada valor recibido.
            valores = [valor.strip() for valor in fila]
            # Obtiene el identificador si existe en la fila o usa una cadena vacía como alternativa.
            orden_id = valores[indice_id] if len(valores) > indice_id else ""
            # Comprueba que la cantidad de valores coincida con la cantidad de encabezados.
            if len(valores) != len(encabezados):
                # Registra la fila completa como error junto con su número e identificador.
                errores.append(
                    (numero, orden_id, "fila", str(fila), "Cantidad de columnas incorrecta")
                )
                # Continúa con la siguiente fila sin procesar el registro incorrecto.
                continue

            # Convierte los encabezados y los valores en un diccionario representativo del registro.
            orden = dict(zip(encabezados, valores))
            # Valida los datos y registra cualquier problema encontrado.
            error = validar_orden(orden, ids_validos)
            # Comprueba si la validación devolvió un error.
            if error:
                # Agrega el detalle del error a la lista correspondiente.
                errores.append((numero, orden_id, *error))
                # Continúa con la siguiente fila sin agregar la orden como válida.
                continue

            # Calcula el estado y los días de atraso de la orden validada.
            estado, atraso = clasificar_orden(
                orden["estado_original"], orden["fecha_programada"], fecha_corte
            )
            # Agrega los datos calculados y la fecha de corte al diccionario de la orden.
            orden.update(estado_calculado=estado, dias_atraso=atraso, fecha_corte=fecha_corte)
            # Incorpora la orden procesada a la lista de órdenes válidas.
            ordenes.append(orden)
            # Guarda el identificador para impedir que aparezca nuevamente.
            ids_validos.add(orden_id)

    # Devuelve por separado las órdenes válidas y los errores registrados.
    return ordenes, errores


def exportar_excel(ruta, ordenes, errores):
    """Crea un archivo Excel con hojas separadas para órdenes válidas y errores."""

    # Crea la carpeta de destino si todavía no existe.
    ruta.parent.mkdir(parents=True, exist_ok=True)
    # Crea un archivo temporal en la misma carpeta para preparar la exportación.
    with NamedTemporaryFile(dir=ruta.parent, suffix=".xlsx", delete=False) as archivo:
        # Convierte el nombre temporal en un objeto de ruta reutilizable.
        temporal = Path(archivo.name)

    # Garantiza la limpieza del archivo temporal incluso si ocurre una excepción.
    try:
        # Abre el archivo temporal como libro Excel usando el motor OpenPyXL.
        with pd.ExcelWriter(temporal, engine="openpyxl") as excel:
            # Escribe la tabla de órdenes válidas en la hoja llamada Ordenes.
            pd.DataFrame(ordenes, columns=COLUMNAS_ORDENES).to_excel(
                excel, sheet_name="Ordenes", index=False
            )
            # Escribe la tabla de errores en la hoja llamada Errores.
            pd.DataFrame(errores, columns=COLUMNAS_ERRORES).to_excel(
                excel, sheet_name="Errores", index=False
            )

            # Recorre cada hoja creada para aplicar formato y opciones de navegación.
            for hoja in excel.sheets.values():
                # Congela la primera fila para mantener visibles los encabezados al desplazarse.
                hoja.freeze_panes = "A2"
                # Activa filtros automáticos sobre el rango completo utilizado por la hoja.
                hoja.auto_filter.ref = hoja.dimensions
                # Recorre las filas de datos a partir de la segunda fila.
                for fila in hoja.iter_rows(min_row=2):
                    # Recorre cada celda de la fila actual.
                    for celda in fila:
                        # Los valores del CSV se muestran como texto, incluso si parecen fórmulas.
                        # Si el valor es texto, fuerza que Excel lo conserve como una cadena.
                        if isinstance(celda.value, str):
                            celda.data_type = "s"
                        # Si el valor es una fecha, aplica el formato de fecha visible esperado.
                        elif isinstance(celda.value, date):
                            celda.number_format = "DD/MM/YYYY"
        # Intenta mover el archivo temporal a la ruta final solicitada.
        try:
            temporal.replace(ruta)
        # Si el destino está bloqueado, transforma el error en un mensaje más orientativo.
        except PermissionError as error:
            raise PermissionError(
                f"No se pudo reemplazar {ruta}. Si está abierto en Excel, cerralo "
                "y volvé a ejecutar el script."
            ) from error
    # Ejecuta siempre la eliminación del archivo temporal al finalizar el proceso.
    finally:
        temporal.unlink(missing_ok=True)


def main():
    """Configura la interfaz de línea de comandos y coordina todo el procesamiento."""

    # Crea el analizador de argumentos usando la descripción del módulo como ayuda.
    parser = argparse.ArgumentParser(description=__doc__)
    # Declara el argumento obligatorio que identifica el CSV de entrada.
    parser.add_argument("--entrada", type=Path, required=True, help="Ruta del CSV de órdenes.")
    # Declara el argumento obligatorio que identifica el archivo Excel de salida.
    parser.add_argument("--salida", type=Path, required=True, help="Ruta del reporte .xlsx.")
    # Declara el argumento obligatorio que recibe la fecha de corte en formato ISO.
    parser.add_argument("--fecha-corte", required=True, help="Fecha de corte en formato AAAA-MM-DD.")
    # Lee y valida los argumentos enviados al ejecutar el script.
    args = parser.parse_args()

    # Intenta convertir el texto recibido en una fecha real.
    try:
        fecha_corte = date.fromisoformat(args.fecha_corte)
    # Termina con un mensaje de uso si la fecha no tiene un formato válido.
    except ValueError:
        parser.error("--fecha-corte debe ser una fecha válida en formato AAAA-MM-DD.")
    # Comprueba que la fecha no haya sido aceptada con una representación alternativa.
    if fecha_corte.isoformat() != args.fecha_corte:
        parser.error("--fecha-corte debe usar el formato AAAA-MM-DD.")
    # Comprueba que la ruta de salida termine con la extensión requerida.
    if args.salida.suffix.lower() != ".xlsx":
        parser.error("--salida debe terminar en .xlsx.")
    # Evita sobrescribir el archivo CSV de entrada al elegir el mismo destino.
    if args.entrada.resolve() == args.salida.resolve():
        parser.error("La entrada y la salida deben ser archivos distintos.")

    # Agrupa el procesamiento y la exportación para manejar sus errores de manera conjunta.
    try:
        # Lee y procesa el CSV para obtener órdenes válidas y registros rechazados.
        ordenes, errores = procesar_csv(args.entrada, fecha_corte)
        # Exporta los resultados al archivo Excel solicitado.
        exportar_excel(args.salida, ordenes, errores)
    # Captura errores de sistema, datos o formato CSV y evita que el programa muestre un rastreo.
    except (OSError, ValueError, csv.Error) as error:
        # Imprime un mensaje claro en la salida de errores.
        print(f"Error: {error}", file=sys.stderr)
        # Devuelve un código distinto de cero para indicar que el proceso falló.
        return 1

    # Cuenta cuántas órdenes válidas quedaron clasificadas como vencidas.
    vencidas = sum(orden["estado_calculado"] == "Vencida" for orden in ordenes)
    
    # Muestra al usuario la fecha de corte utilizada en el procesamiento.
    print(f"Fecha de corte: {fecha_corte:%d/%m/%Y}")
    # Muestra un resumen de las órdenes válidas, vencidas y rechazadas.
    print(f"Órdenes válidas: {len(ordenes)} | Vencidas: {vencidas} | Rechazadas: {len(errores)}")
    # Muestra la ruta absoluta del archivo Excel generado.
    print(f"Archivo generado: {args.salida.resolve()}")
    # Devuelve cero para indicar que el procesamiento finalizó correctamente.
    return 0


# Ejecuta main únicamente cuando el archivo se inicia como programa principal.
if __name__ == "__main__":
    # Convierte el código de salida de main en el estado de terminación del proceso.
    raise SystemExit(main())
