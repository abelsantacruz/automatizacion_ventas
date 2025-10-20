/ Descripción del script
	
	Este script automatiza la consolidación y análisis de archivos .xlsx de ventas distribuidos por diferentes regiones y fechas
	
/ Incluye una validación de clasificación con lógica aplicable para dos escenarios comunes:
	
	- Por nombre 	: Se utiliza el nombre del archivo en el formato esperado ej: (sales_region_month_year.xlsx) para determinar la fecha y región
	- Por contenido : Si el nombre no sigue el formato esperado, se valida leyendo la columna "Date" alojado dentro del archivo

/ También posee dos modos de operación:

	- copy : copia los archivos desde Input/ hacia Output/ conservando los archivos originales (bueno para hacer debug y testear el script)
	- move : mueve los archivos desde Input/ hacia Output/ eliminando los originales, se solicita confirmación antes de eliminar

/ Los archivos generados son comprendidos por una carpeta mensual que contiene un Excel consolidado con dos hojas:

	Datos Consolidados:Contiene todos los registros del mes, incluyendo columnas:
	Date, Region, Salesperson, Product, Quantity, UnitPrice, Total, RegionOrigen.

	Y Ranking Productos: Ranking de productos más vendidos del mes, ordenado por Quantity y Total.

/ Como registro de ejecución el script genera automáticamente un log con fecha dentro de la carpeta Output/ incluyendo:
	Modo de ejecución : (copy o move)
	Cantidad de archivos procesados y con errores
	Duplicados detectados
	Advertencias o inconsistencias
	
/ Se implementaron códigos de error con comportamientos esperados por el script	
	Código		Descripción
	E001		Faltan columnas esperadas en el .xlsx			
	E002		Datos inválidos, ya sea valores no númericos o fechas incorrectas
	E003		Error al leer el archivo, se sospechará que el archivo esté dañado
	E004		Nombre de archivo no reconocido o sin información de fecha
	
/ Requisitos 
 
	Python 3.8 o superior
	El script depende de las librerías pandas & openpyxl
	
	Para instalarlas ejecuta en la terminal "pip install pandas openpyxl" 

/ Uso del script

	- Colocar el script en la misma carpeta donde se encuentren los archivos .xlsx a procesar
	- Se ejecuta el script desde cualquier terminal eligiendo el modo de operación como argumento --copy o --move 
		ej. "python .\automatizacion_ventas.py --move"

/ Recomendaciones 
	
	- Mantener la carpeta Input/ solo con archivos nuevos para evitar duplicados indeseados
	- Asegurarse que los nombres de los archivos sigan el formato estandarizado si es posible
	- Comprobar si los archivos poseen valores legítimos y precisos
	- Revisar los logs luego de cada ejecución

	
	