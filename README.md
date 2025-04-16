
# Cotizaciones BCRA a Excel 🇦🇷📈

Este proyecto consulta la API del Banco Central de la República Argentina (BCRA) para obtener cotizaciones diarias de monedas extranjeras, y guarda el valor del dólar y del euro en una hoja de Excel.

## 🚀 Funcionalidad

- Consulta la API oficial del BCRA.
- Filtra las monedas deseadas (USD y EUR).
- Escribe las cotizaciones actuales en celdas específicas de un archivo Excel (`cotizaciones.xlsx`).

## 📦 Requisitos

- Python 3.8 o superior
- Excel instalado en el sistema (requerido por `xlwings`)

## 🛠️ Instalación

1. Cloná el repositorio:

```bash
git clone https://github.com/tu-usuario/cotizaciones-bcra-excel.git
cd cotizaciones-bcra-excel
pip install xlwings requests
```
## 📁 Estructura del código
- Request_Api: Clase para conectarse a la API del BCRA y filtrar resultados.
- Escribir_excel: Clase para escribir los valores obtenidos en el Excel.