## Controlador de Excel
Controlador de Excel es una aplicación de escritorio simple, desarrollada en Python utilizando tkinter, que permite a los usuarios gestionar y registrar datos de equipos en una hoja de cálculo de Excel. Los usuarios pueden ingresar información como la referencia del equipo, SKU, RMA, precio y trabajo realizado, y guardar estos datos en un archivo Excel. La aplicación también permite crear una carpeta basada en el nombre de la referencia del equipo, donde se pueden almacenar imágenes o archivos relacionados.

# Características
Ingreso de datos para Equipo, SKU, RMA, Precio y Trabajo Realizado.
Guarda los datos ingresados en un archivo Excel llamado Salida Equipos.xlsx.
Crea y abre automáticamente una carpeta con el nombre de la referencia del equipo para almacenar imágenes o archivos relacionados.
Limita la cantidad de entradas a un máximo de 36 equipos.
Previene entradas duplicadas verificando si la referencia del equipo ya existe en el archivo Excel.
Requisitos

# Instala las dependencias necesarias ejecutando:

pip install openpyxl
pip install tkinter

# Uso
Ingresar información: Ingresa el nombre del equipo en el campo Equipo y presiona "Buscar Equipo" para desbloquear los otros campos.
Guardar datos: Rellena todos los campos y presiona el botón "Guardar" para almacenar los datos en el archivo Excel.
Crear carpeta: Presiona el botón "Guardar imagen" para crear una carpeta con el nombre de la referencia del equipo y abrirla automáticamente.


# Estructura del Proyecto

ExcelController/
│
├── ExcelController.py   # Código principal de la aplicación
└── README.md            # Instrucciones y guía del proyecto
