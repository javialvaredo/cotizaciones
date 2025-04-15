import requests
from datetime import datetime
import xlwings as xw  

def convertir_formato_fecha(fecha):
    fecha_iso = fecha
    fecha_obj = datetime.fromisoformat(fecha_iso.replace('Z', '+00:00'))
    return fecha_obj.strftime('%d-%m-%y')

def monedas_mayorista(moneda):
    response = requests.get(moneda)
    if response.status_code == 200:
        data = response.json()
        tipo_moneda = data['moneda']
        valor_vendedor = data['venta']
        fecha_cierre = data['fechaActualizacion']
        fecha_con_formato = convertir_formato_fecha(fecha_cierre)
        return (tipo_moneda, fecha_con_formato, valor_vendedor)
    else:
        print("Error al obtener los datos.")
        return None

# Conectar al libro abierto de Excel
def escribir_valor_excel(moneda, valor, nombre_libro):
    if moneda[0] == "dolar":
        celda='B2'
    else:
        celda='B3'
    
    wb = xw.Book(nombre_libro)  # Asegurate de que el archivo esté abierto y este sea el nombre correcto
    hoja = wb.sheets[0]  # Podés cambiar por el nombre: wb.sheets['Hoja1']
    hoja.range(celda).value = valor

if __name__ == '__main__':

    dolar = ("dolar","https://dolarapi.com/v1/dolares/mayorista")
    euro = ("euro", "https://dolarapi.com/v1/cotizaciones/eur")
    file = r"C:\Users\jalvaredo\OneDrive - CV CONTROL SA\Sincro\Bancos\cotizaciones.xlsx"

    datos = monedas_mayorista(dolar[1])
    if datos:
        print(f"Cierre vendedor de {datos[0]} divisa del {datos[1]}: $ {datos[2]}")
        escribir_valor_excel(dolar, datos[2], file )  # 👉 Esto actualiza la celda B2 en tiempo real

    datos = monedas_mayorista(euro[1])
    if datos:
        print(f"Cierre vendedor de {datos[0]} divisa del {datos[1]}: $ {datos[2]}")
        escribir_valor_excel(euro, datos[2], file) 
