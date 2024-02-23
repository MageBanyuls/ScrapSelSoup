from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import time
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.support.ui import Select
import pandas as pd


archivo_excel=r"C:\Users\USUARIO\Downloads\Excel_agos - copia.xlsx"

def login_sii (driver,rut,clave):


        try:

            driver.get("https://zeusr.sii.cl//AUT2000/InicioAutenticacion/IngresoRutClave.html?https://misiir.sii.cl/cgi_misii/siihome.cgi")
    
            ruter_input =  driver.find_element(By.ID, "rutcntr")
            ruter_input.send_keys(rut)
    
            pass_input = driver.find_element(By.ID, "clave")
            pass_input.send_keys(clave)
    
            btn_ingreso = driver.find_element(By.ID, "bt_ingresar")
            btn_ingreso.click()

            time.sleep(2)
    
            try:
                alert = driver.switch_to.alert
                alert.dismiss()
            except:
                pass
    
            try:
                alert = driver.switch_to.alert
                alert.dismiss()
            except:
                pass
    
            try:
                driver.find_element(By.ID, "titulo")
                print("no se pudo hacer login")
                return False
    
            except NoSuchElementException:
                print('login exitosos')
                #aca analizo si esta la pantalla de siguiente"
                try:
                    try:
                        modal = driver.find_element(By.CSS_SELECTOR, 'div.modal-dialog')
    
                        if modal:
                            # Si hay un modal, hacer clic en el botón de cierre
                            btn_cierre_modal = driver.find_element(By.XPATH, '//*[@id="ModalEmergente"]/div/div/div[3]/button')
                            btn_cierre_modal.click()
                    except:
                        pass
    
                    time.sleep(2)
    
                    try:
                        modal = driver.find_element(By.ID,'myMainCorreoVigente')
                        if modal.is_displayed():
                            driver.execute_script("arguments[0].style.display = 'none';", modal)
    
                    except:
                        pass
    
    
                except NoSuchElementException:
                    boton_siguiente = driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/p[2]/a[1]')
                    boton_siguiente.click()
    
                    try:
                        try:
                            alert = driver.switch_to.alert
                            alert.dismiss()
                        except:
                            pass
    
                        try:
                            alert = driver.switch_to.alert
                            alert.dismiss()
                        except:
                            pass
    
                        try:
                            modal = driver.find_element(By.CSS_SELECTOR, 'div.modal-dialog')
    
                            if modal:
                                # Si hay un modal, hacer clic en el botón de cierre
                                btn_cierre_modal = driver.find_element(By.XPATH, '//*[@id="ModalEmergente"]/div/div/div[3]/button')
                                btn_cierre_modal.click()
                        except:
                            pass
    
                        time.sleep(2)
    
                        try:
                            modal = driver.find_element(By.ID,'myMainCorreoVigente')
                            if modal.is_displayed():
                                driver.execute_script("arguments[0].style.display = 'none';", modal)
                        except:
                            pass
    
                    except NoSuchElementException:
                        print('keseste erroooor!!')
                return True
    
        except TimeoutException as timeout_error:
    
            print(f"Error de tiempo de espera para RUT {rut}: {timeout_error}")
        except Exception as e:
            print(f"Error durante el proceso para RUT {rut}: {e}")



def getData(parte_df):

    options = Options()
    
    
    #options.add_argument("--headless")
    driver_path = r"C:\Users\USUARIO\Downloads\chromedriver-win64\chromedriver.exe"
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    for index, row in parte_df.iterrows():
        rut=row['Rut']
        clave=row['Clave SII']

        login_success=login_sii(driver,rut,clave)
        time.sleep(1)
    
        #busqueda facturas de consulta seguimiento

        wait=WebDriverWait(driver,20)

        driver.get("https://www4.sii.cl/djconsultarentaui/internet/#/agenteretenedor/")

        select_anio = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="periodo"]')))

        select_anio.send_keys('2023')
        time.sleep(1)

        boton_aceptar = driver.find_element(By.XPATH,'//*[@id="my-wrapper"]/div[2]/div/div/div/div/div[3]/div/div/div[2]/form/button')
        boton_aceptar.click()
        time.sleep(1)

        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="up"]/div[5]')))
        # Obtener el html con selenium
        html_text = driver.page_source

        # Crear BSoup con el Html
        soup = BeautifulSoup(html_text, "html.parser")

        # Encontrar todos los divs que empicen "div-resumen-"
        divs_resumen = soup.find_all(lambda tag: tag.name == "div" and tag.get("id") and tag.get("id").startswith("div-resumen-"))

        # Crear un diccionario para almacenar los datos
        datos = {}

        # Iterar sobre cada div de resumen
        for resumen in divs_resumen:
            # Buscar el div con la clase "panel-body" dentro del div de resumen
            panel_body = resumen.find("div", class_="panel-body")
            
            # Si no encuentro el panel body, omitimos el resumen
            if not panel_body:
                continue
            
            # Encontrar la tabla dentro del panel-body
            tabla = panel_body.find("table")
            
            # Si no se encuentra la tabla, omitir este resumen
            if not tabla:
                continue
            
            # Encontrar todas las filas de datos en la tabla
            filas_tabla = tabla.find_all("tr")
            
            # Obtener los nombres de las columnas de la penúltima fila porque en la ultima estan los valores y las anteriores no sirven
            columnas = [td.text.strip() for td in filas_tabla[-2].find_all("td")]
            
            # Obtener los valores de la última fila
            valores = [td.text.strip() for td in filas_tabla[-1].find_all("td")]
            
            # Actualizar el nombre de las columnas con el código del cliente
            codigo_cliente = resumen.find("span", class_="bold-row").text.strip()
            columnas = [f"{codigo_cliente}-{columna}" for columna in columnas]
            
            # Agregar los valores al diccionario de datos
            datos[codigo_cliente] = dict(zip(columnas, valores))

            # Iterar sobre cada clave y valor en el diccionario datos
            for codigo_cliente, valores_cliente in datos.items():
                # Iterar sobre cada nombre de columna y valor en los valores del cliente
                for nombre_columna, valor in valores_cliente.items():
                    # Verificar si la columna ya existe en el DataFrame filtrado
                    if nombre_columna not in df.columns:
                        # Si la columna no existe, agregarla al DataFrame filtrado con valores NaN
                        df[nombre_columna] = float('nan')
                    
                    # Asignar el valor del cliente a la columna correspondiente en el DataFrame filtrado
                    df.at[index, nombre_columna] = valor

                    

                    time.sleep(2)


                df.at[index, 'Revisado'] = 'Si'
            
            print(datos)

        


        if login_success:
            df.at[index, 'Revisado'] = 'Si'
        else:
            df.at[index, 'Revisado'] = 'Contraseña'
        df.to_excel(archivo_excel, index=False)
    
    driver.quit()


        










if __name__ == "__main__":
 
    # Leer el archivo Excel con los clientes a consultar
    df = pd.read_excel(archivo_excel)

    # Crear un DataFrame para almacenar los resultados
    resultados_df = pd.DataFrame()

    df_no = df[df['Revisado'] == 'No']
    # Dividir el DataFrame en partes para distribuir entre los trabajadores
    num_threads = 20
    particiones = np.array_split(df_no, num_threads)


    # Configurar el número máximo de trabajadores
    max_workers = 20

    # Crear un ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=max_workers) as executor:

        futures = []

        # Lanzar una tarea para cada porción del DataFrame
        for parte_df in particiones:
            future = executor.submit(getData, parte_df)
            futures.append(future)

        # Esperar a que todas las tareas se completen
        for future in futures:
            future.result()
