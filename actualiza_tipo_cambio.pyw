import requests
import xlwings as xw  


class Request_Api:
    def __init__(self, url):
        self.url = url
        self.headers = {'Accept-Language': 'es-AR'}
        self.response = None

    def get_request(self):
        try:
            self.response = requests.get(self.url, headers=self.headers, verify=False)
            self.response.raise_for_status()  # Will raise an HTTPError if the HTTP request returned an unsuccessful status code
            self.response = self.response.json()  # Parse the JSON response
        except requests.exceptions.RequestException as e:
            print(f'Error al obtener datos: {e}')
            return None
        return self.response
    
    def filtrar_request(self, monedas, monedas_deseadas):
        euro = None
        dolar = None

        for moneda in monedas:
            if moneda['codigoMoneda'] in monedas_deseadas:
                
                if moneda['codigoMoneda'] == 'EUR':
                    euro = moneda['tipoCotizacion']

                elif moneda['codigoMoneda'] == 'USD':
                    dolar = moneda['tipoCotizacion']
        return (dolar, euro)
   
    
class Escribir_excel:
    def escribir_valor_excel(self, resultado, nombre_libro, nombre_hoja):
        
        wb = xw.Book(nombre_libro)  # Asegurate de que el archivo esté abierto y este sea el nombre correcto
        hoja = wb.sheets[nombre_hoja]  # Podés cambiar por el nombre: wb.sheets['Hoja1']
        hoja.range('D2').value = resultado[0]
        hoja.range('F2').value = resultado[1]



def main():
    url = 'https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones'

    nombre_libro = r"C:\Users\jalvaredo\OneDrive - CV CONTROL SA\Sincro\Bancos\BANCOS actualizado.xlsm"
    nombre_hoja = "CASHFLOWCV"

    request = Request_Api(url)
    datos = request.get_request()
    if not datos or 'results' not in datos:
        print("No se pudo obtener información de la API.")
        return

    monedas = datos['results']['detalle']  # extraer lista de monedas
    monedas_deseadas = ['USD', 'EUR']  #Filtrar lista de monedas que interesan

    resultado = request.filtrar_request(monedas, monedas_deseadas)
    
    print(f"Dolar: {resultado[0]} -  Euro: {resultado[1]}")

    actualizar_excel = Escribir_excel()
    actualizar_excel.escribir_valor_excel(resultado, nombre_libro, nombre_hoja)



if __name__ == '__main__':
    main()
    
