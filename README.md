Gestión de Empresas con Clasificación ARL en Excel
Descripción
Este es un programa de consola en Python diseñado para ayudar en la gestión de un listado de empresas, enfocándose en la recopilación y organización de datos relevantes para la clasificación de Riesgos Laborales (ARL) en Colombia. La información de las empresas se almacena persistentemente en un archivo Excel (.xlsx), lo que facilita su consulta y edición.

La herramienta permite a los usuarios:

Añadir nuevas empresas de forma interactiva, guiando al usuario campo por campo y realizando validaciones en tiempo real para asegurar la calidad de los datos.

Actualizar la información de empresas existentes, buscando por su Razón Social y permitiendo modificar campos específicos.

Realizar cargas masivas de empresas pegando datos estructurados directamente en la consola, ideal para migrar grandes volúmenes de información o automatizar la adición de múltiples registros.

El script incluye validaciones robustas para campos como fechas, correos electrónicos, teléfonos, números de NIT/identificación, ingresos y códigos CIIU, asegurando que los datos cumplan con formatos esperados y reglas de negocio específicas (ej. un NIT es "N/A" si el Tipo de Identificación es "NO NIT"). Además, aplica estilos automáticos al archivo Excel para mejorar su legibilidad.

Funcionalidades Destacadas
Almacenamiento en Excel: Utiliza openpyxl para manejar la lectura y escritura de datos en el archivo Listado_Empresas_ARL_Automatizado.xlsx.

Interfaz de Consola Intuitiva: Ofrece un menú claro y solicita la información de manera guiada.

Validación de Datos en Tiempo Real: Previene errores comunes de formato en campos clave como FECHA_DE_MATRICULA, TELEFONO, CORREO, NUMERO_DE_NIT y CIIU.

Lógica de Negocio Integrada: Gestiona la relación entre TIPO_DE_IDENTIFICACION y NUMERO_DE_NIT.

Formato de Carga Masiva Estandarizado: Facilita la importación de datos desde otras fuentes mediante un formato pipe-separated (|).

Estilización Automática de Excel: Mejora la presentación visual de la hoja de cálculo generada.

Cómo Empezar
Requisitos
Asegúrate de tener Python instalado en tu sistema. Este script ha sido desarrollado y probado con Python 3.x.

Las librerías externas necesarias son:

openpyxl

Puedes instalarlas usando pip:

Bash

pip install openpyxl
Ejecución
Clona este repositorio o descarga los archivos Python_Empresas.py y el README.md.

Abre tu terminal o símbolo del sistema.

Navega hasta el directorio donde guardaste el script.

Ejecuta el programa con el siguiente comando:

Bash

python Python_Empresas.py
El script creará o cargará el archivo Listado_Empresas_ARL_Automatizado.xlsx y te presentará el menú principal.

Uso
Sigue las opciones del menú en la consola:

Agregar nueva empresa (manual): Ingresa los datos solicitados paso a paso.

Actualizar empresa existente (por Razón Social): Busca una empresa por su nombre y selecciona los campos a modificar.

Cargar múltiples empresas (desde consola): Pega tus datos en el formato CAMPO1|CAMPO2|...|CAMPO15 (consulta los ENCABEZADOS en el código para el orden exacto). Finaliza la carga escribiendo FIN_CARGA.

Salir y Guardar: Guarda todos los cambios en el archivo Excel. Asegúrate de que el archivo no esté abierto en otra aplicación (como Microsoft Excel) al momento de guardar para evitar errores de permisos.

Estructura del Proyecto
.
├── Python_Empresas.py       # Script principal del programa
└── README.md                # Este archivo de documentación
└── Listado_Empresas_ARL_Automatizado.xlsx # Archivo Excel generado/usado por el script
Contribuciones
Las contribuciones son bienvenidas. Si tienes ideas para mejorar, informes de errores o quieres añadir nuevas funcionalidades, no dudes en abrir un issue o enviar un pull request.

Licencia
Este proyecto está bajo la Licencia MIT. Consulta el archivo LICENSE para más detalles.
