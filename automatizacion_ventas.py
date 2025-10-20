#IMPORTANTE, se necesita de la liberia pandas y la de openpyxl para ejecutar el script

from pathlib import Path
from datetime import datetime
import pandas as pd
import shutil
import logging
import re
import sys
import argparse
from typing import Dict, Optional, Tuple, List

# columnas en orden esperado que deben estar presentes en cada archivo excel
COLUMNAS_ESPERADAS = {"Date", "Region", "Salesperson", "Product", "Quantity", "UnitPrice"}

# uso de expresion regular para validar y extraer los datos del nombre del archivo, formato esperado = sales_<region>_<month>_<year>.xlsx
PATRON_ARCHIVO = re.compile(r"^sales_([a-zA-Z]+)_([a-zA-Z]+)_([0-9]{4})\.xlsx$", re.IGNORECASE)

# codigos de eror utilizados en el logger para catalogar los errores
E001, E002, E003, E004 = 'E001', 'E002', 'E003', 'E004'

# conf. de diccionario que traduce el mes en ingles en numero y nombre en espaniol
MESES = {
    'january': (1, 'Enero'), 'february': (2, 'Febrero'), 'march': (3, 'Marzo'),
    'april': (4, 'Abril'), 'may': (5, 'Mayo'), 'june': (6, 'Junio'),
    'july': (7, 'Julio'), 'august': (8, 'Agosto'), 'september': (9, 'Septiembre'),
    'october': (10, 'Octubre'), 'november': (11, 'Noviembre'), 'december': (12, 'Diciembre')
}

# extraccion de datos1 desde nombre del archivo
def analizar_nombre_archivo(nombre_archivo: str) -> Optional[Dict]:

    # se utiliza la expresion regular para buscar coincidencias y extraer los datos
    coincidencia = PATRON_ARCHIVO.search(nombre_archivo)
    if not coincidencia:
        return None
    region, mes_texto, anio = coincidencia.group(1), coincidencia.group(2).lower(), int(coincidencia.group(3))
    if mes_texto not in MESES:
        return None
    numero_mes, nombre_mes = MESES[mes_texto]

    # luego de validar se retorna un dicc. con la info extraida y estructurada
    return {'region': region, 'mes_texto': mes_texto, 'numero_mes': numero_mes, 'nombre_mes': nombre_mes, 'anio': anio}


# extraccion de datos2 desde contenido del archivo
def obtener_info_desde_excel(ruta_archivo: Path) -> Optional[Dict]:
    try:

        # se intenta leer las columnas Date y Region del .xls usando openpyxl
        df = pd.read_excel(ruta_archivo, usecols=["Date", "Region"], engine="openpyxl")

        # se convierte la columna Date a tipo datetime y se limpian los valores invalidos
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        if df.empty:
            return None

        # obtiene el mes y anio mas frecuentes en caso de inconsistencias con los archivos entrantes
        mes = int(df["Date"].dt.month.mode()[0])
        anio = int(df["Date"].dt.year.mode()[0])

        # se obtiene la region si existe o se marca como Desconocida
        region = str(df["Region"].iloc[0]) if "Region" in df.columns else "Desconocida"
        nombre_mes = list(MESES.values())[mes - 1][1]

        # traduce el mes numerico a espaniol y retorna los datos estructurados
        return {"region": region, "numero_mes": mes, "nombre_mes": nombre_mes, "anio": anio}
    except Exception:
        return None


# creacion de la carpeta de salida/output
def crear_carpeta_output(base_output: Path, anio: int, numero_mes: int, nombre_mes: str) -> Path:
    carpeta = base_output / str(anio) / f"{numero_mes:02d}_{nombre_mes}"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


# leer .xls y si no se puede leer se arroja error correspondiente
def leer_excel(ruta_archivo: Path, info: Dict, logger: logging.Logger) -> pd.DataFrame:
    try:
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"[{E003}] No se pudo leer el archivo: {e}")

    # se normalizan los nombres de columnas stripeando los espacios
    df.columns = [str(c).strip() for c in df.columns]

    # verifica que las columnas esperadas esten presentes
    faltantes = COLUMNAS_ESPERADAS - set(df.columns)
    if faltantes:
        raise ValueError(f"[{E001}] Faltan columnas requeridas: {sorted(faltantes)}")
    
    # conversion de tipos 
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce")
    df["Date"] = pd.to_datetime(df["Date"].astype(str), dayfirst=True, errors="coerce")

    # validacion de datos invalidos
    errores = []
    if df["Date"].isna().any(): errores.append('fecha invalidas en "Date"')
    if df["Quantity"].isna().any(): errores.append('valores no numericos en "Quantity"')
    if df["UnitPrice"].isna().any(): errores.append('valores no numericos en "UnitPrice"')

    # si hay errores se arroja la excepcion de error correspondiente
    if errores:
        raise ValueError(f"[{E002}] Datos invalidos: {'; '.join(errores)}")

    # validacion de consistencia del mes/anio del dataframe
    if not df["Date"].empty:
        mes_real = int(df["Date"].dt.month.mode()[0])
        anio_real = int(df["Date"].dt.year.mode()[0])
        if mes_real != info["numero_mes"] or anio_real != info["anio"]:
            logger.warning(
                "Inconsistencia detectada: %s contiene datos de %02d/%d (esperado: %02d/%d)",
                ruta_archivo.name, mes_real, anio_real, info["numero_mes"], info["anio"]
            )

    # se retorna el dataframe validado
    return df


# procesamiento y validacion de los archivos mensuales
def procesar_mes(lista_archivos_mes: List[tuple], carpeta_destino: Path, logger: logging.Logger, contadores: Dict[str, int]):
    dataframes = []
    for archivo, info in lista_archivos_mes:
        try:
            df = leer_excel(archivo, info, logger)
        except ValueError as e:
            # se capturan los valores de error y se actualizan los contadores
            mensaje = str(e)
            if "[" in mensaje:
                codigo = mensaje.split("]")[0].strip("[")
                contadores[codigo] = contadores.get(codigo, 0) + 1
            contadores["errores"] += 1
            logger.error("Archivo omitido: %s — %s", archivo.name, e)
            continue
        except Exception as e:
            contadores["errores"] += 1
            logger.error("Error inesperado en %s: %s", archivo.name, e)
            # se prevene que el proceso acabe debido a un error inesperado
            continue
        
        # se agrega la informacion adicional, se calcula el total y se guarda el df
        df["RegionOrigen"] = info["region"]
        df["Total"] = df["Quantity"] * df["UnitPrice"]
        dataframes.append(df)
        contadores["archivos_ok"] += 1
        
        # si falla en obtener datos validos para el df se finaliza
    if not dataframes:
        logger.info("No hay datos validos para %s", carpeta_destino)
        return

    # concatenan los df ya ordenados
    consolidado = pd.concat(dataframes, ignore_index=True)
    consolidado["Date"] = consolidado["Date"].dt.strftime("%d/%m/%Y")
    consolidado = consolidado.sort_values(by="Date")

    # se crea el ranking de productos
    ranking = (
        consolidado.groupby("Product", as_index=False)
        .agg({"Quantity": "sum", "Total": "sum"})
        .sort_values(["Quantity", "Total"], ascending=False)
    )

    # generacion del archivo .xls
    anio = carpeta_destino.parent.name
    codigo_mes = carpeta_destino.name.split("_")[0]
    ruta_reporte = carpeta_destino / f"Ventas_Consolidadas_{anio}_{codigo_mes}.xlsx"

    with pd.ExcelWriter(ruta_reporte, engine="openpyxl") as writer:
        consolidado.to_excel(writer, index=False, sheet_name="Datos_Consolidados")
        ranking.to_excel(writer, index=False, sheet_name="Ranking_Productos")
    
    # log del reporte generado
    logger.info("Reporte generado: %s (%d registros, %d productos)", ruta_reporte, len(consolidado), len(ranking))
    contadores["reportes_generados"] += 1

# conf. del logger
def configurar_logger(carpeta_output: Path) -> logging.Logger:
    # asegura que la carpeta de salida/output exista
    carpeta_output.mkdir(parents=True, exist_ok=True)
    fecha = datetime.now().strftime('%Y-%m-%d')
    ruta_log = carpeta_output / f'log_{fecha}.log'

    logger = logging.getLogger('automatizacion_ventas')
    logger.setLevel(logging.INFO)
    formato = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    # se evita agregar handlers si la funcion se llama varias veces
    if not logger.handlers:
        archivo_handler = logging.FileHandler(ruta_log, encoding='utf-8')
        archivo_handler.setFormatter(formato)
        consola_handler = logging.StreamHandler(sys.stdout)
        consola_handler.setFormatter(formato)
        logger.addHandler(archivo_handler)
        logger.addHandler(consola_handler)
    # se apunta a donde queda el archivo del log
    logger.info("Logger inicializado. Archivo de registro: %s", ruta_log)
    return logger

# main 
def main():
    
    # conf. de los args
    parser = argparse.ArgumentParser(description="Automatizacion de consolidacion de ventas")
    grupo_modo = parser.add_mutually_exclusive_group()
    grupo_modo.add_argument("--copy", action="store_true", help="Copia los archivos (modo por defecto)")
    grupo_modo.add_argument("--move", action="store_true", help="Mueve los archivos (requiere confirmación)")
    args = parser.parse_args()

    # se determina el modo de operacion, por default es copy
    modo = "copy"
    if args.move:
        confirmacion = input("Estas a punto de mover los archivos originales.\n¿Desea continuar? (S/N): ").strip().lower()
        if confirmacion != "s":
            print("Operacion cancelada por el usuario.")
            return
        modo = "move"

    # conf. del path del archivo, por defecto en la misma ubicacion de la entrada/input
    ruta_base = Path(".")
    carpeta_entrada = ruta_base / "Input"
    carpeta_output_base = ruta_base / "Output"
    logger = configurar_logger(carpeta_output_base)
    logger.info("Modo de ejecucion: %s", modo.upper())
    
    # se valida si existe la carpeta de entrada y si existen archivos .xlsx
    if not carpeta_entrada.exists():
        logger.error("No se encontro la carpeta Input/.")
        return

    archivos_entrada = [f for f in carpeta_entrada.glob("*.xlsx") if f.is_file()]
    if not archivos_entrada:
        logger.warning("No se encontraron archivos .xlsx en Input/.")
        return

    # agrupacion de archivos por mes y anio
    grupos_mensuales: Dict[Tuple[int, int, str], List[tuple]] = {}
    contadores = {"archivos_ok": 0, "errores": 0, "duplicados": 0, "reportes_generados": 0,
                  E001: 0, E002: 0, E003: 0, E004: 0}

    for archivo in archivos_entrada:
        info = analizar_nombre_archivo(archivo.name)
        # si no se cumple con el patron se intenta inferir con el contenido desde el archivo
        if not info:
            info = obtener_info_desde_excel(archivo)
            # se levanta una advertencia si se infiere la info desde el contenido y cuando no se puede determinar mes y anio
            if info:
                logger.warning("[%s] Nombre invalido: %s — inferido desde contenido (%02d/%d)",
                               E004, archivo.name, info["numero_mes"], info["anio"])
            else:
                logger.error("[%s] No se pudo determinar mes/anio: %s", E004, archivo.name)
                contadores["errores"] += 1
                contadores[E004] += 1
                continue
        
        # creacion de clave para agrupar los archivos en el dict
        clave = (info["anio"], info["numero_mes"], info["nombre_mes"])
        grupos_mensuales.setdefault(clave, []).append((archivo, info))

    # procesamiento de los grupos mensuales
    for (anio, numero_mes, nombre_mes), archivos_mes in grupos_mensuales.items():
        carpeta_destino = crear_carpeta_output(carpeta_output_base, anio, numero_mes, nombre_mes)
        archivos_procesados = []
        for archivo, info in archivos_mes:
            destino = carpeta_destino / archivo.name
            # procesamiento de duplicados
            if destino.exists():
                timestamp = datetime.now().strftime("%m%d%H%M")
                destino = carpeta_destino / f"{archivo.stem}_dup_{timestamp}{archivo.suffix}"
                contadores["duplicados"] += 1
                logger.warning("Archivo duplicado: %s → Guardado como %s", archivo.name, destino.name)

            # operacion de movimiento o copia
            if modo == "copy":
                shutil.copy2(str(archivo), str(destino))
            else:
                shutil.move(str(archivo), str(destino))
            archivos_procesados.append((destino, info))

        procesar_mes(archivos_procesados, carpeta_destino, logger, contadores)

    # logger resumen
    logger.info("\n========== RESUMEN =========")
    logger.info("Modo: %s", modo.upper())
    logger.info("Archivos procesados correctamente: %d", contadores["archivos_ok"])
    logger.info("Duplicados: %d", contadores["duplicados"])
    logger.info("Reportes generados: %d", contadores["reportes_generados"])
    logger.info("Archivos con errores: %d", contadores["errores"])
    for codigo in [E001, E002, E003, E004]:
        if contadores[codigo] > 0:
            logger.info("  %s → %d incidencias", codigo, contadores[codigo])
    logger.info("================================\n")


if __name__ == "__main__":
    main()
