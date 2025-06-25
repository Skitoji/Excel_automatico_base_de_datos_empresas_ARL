# Gestión de Empresas con Clasificación ARL en Excel

## 📝 Descripción

Este es un programa de consola en Python diseñado para ayudar en la gestión de un listado de empresas, enfocándose en la recopilación y organización de datos relevantes para la clasificación de Riesgos Laborales (ARL) en Colombia. La información de las empresas se almacena persistentemente en un archivo Excel (`.xlsx`), lo que facilita su consulta y edición.

La herramienta permite a los usuarios:

* **Añadir nuevas empresas** de forma interactiva, guiando al usuario campo por campo y realizando validaciones en tiempo real para asegurar la calidad de los datos.
* **Actualizar la información** de empresas existentes, buscando por su Razón Social y permitiendo modificar campos específicos.
* Realizar **cargas masivas de empresas** pegando datos estructurados directamente en la consola, ideal para migrar grandes volúmenes de información o automatizar la adición de múltiples registros.

El script incluye **validaciones robustas** para campos como fechas, correos electrónicos, teléfonos, números de NIT/identificación, ingresos y códigos CIIU, asegurando que los datos cumplan con formatos esperados y reglas de negocio específicas (ej. un NIT es "N/A" si el Tipo de Identificación es "NO NIT"). Además, aplica estilos automáticos al archivo Excel para mejorar su legibilidad.

## ✨ Funcionalidades Destacadas

* **Almacenamiento en Excel:** Utiliza `openpyxl` para manejar la lectura y escritura de datos en el archivo `Listado_Empresas_ARL_Automatizado.xlsx`.
* **Interfaz de Consola Intuitiva:** Ofrece un menú claro y solicita la información de manera guiada.
* **Validación de Datos en Tiempo Real:** Previene errores comunes de formato en campos clave como `FECHA_DE_MATRICULA`, `TELEFONO`, `CORREO`, `NUMERO_DE_NIT` y `CIIU`.
* **Lógica de Negocio Integrada:** Gestiona la relación entre `TIPO_DE_IDENTIFICACION` y `NUMERO_DE_NIT`.
* **Formato de Carga Masiva Estandarizado:** Facilita la importación de datos desde otras fuentes mediante un formato pipe-separated (`|`).
* **Estilización Automática de Excel:** Mejora la presentación visual de la hoja de cálculo generada.

## 🚀 Cómo Empezar

### Requisitos

Asegúrate de tener [Python](https://www.python.org/downloads/) instalado en tu sistema (versión 3.x recomendada).

Las librerías externas necesarias son:
* `openpyxl`

Puedes instalarlas usando `pip`:

```bash
pip install openpyxl
