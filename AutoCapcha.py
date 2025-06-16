import time
import pytesseract
import pandas as pd
import cv2
import numpy as np
from PIL import Image
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# Configurar la ruta a tesseract si es necesario
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Lee el archivo Excel
df = pd.read_excel("datos.xlsx")  # Asegúrate de tener una columna 'documento'

# Inicializa el navegador
driver = webdriver.Chrome()
driver.get("https://www.miseguridadsocial.gov.co/")  # Coloca la URL real del formulario

for index, row in df.iterrows():
    documento = str(row['documento'])

    # Selecciona tipo de documento
    Select(driver.find_element(By.ID, "ddlTipoDocumento")).select_by_visible_text("Tarjeta de Identidad")

    # Ingresa número de documento
    campo_numero = driver.find_element(By.ID, "txtNumeroDocumento")
    campo_numero.clear()
    campo_numero.send_keys(documento)

    # Obtiene el captcha como imagen
    captcha_element = driver.find_element(By.ID, "imgCaptcha")  # Ajusta el ID según la página
    captcha_bytes = captcha_element.screenshot_as_png
    image = Image.open(BytesIO(captcha_bytes))

    # Limpieza básica de imagen con OpenCV
    img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    _, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)

    # OCR para leer el texto del CAPTCHA
    captcha_text = pytesseract.image_to_string(img_bin, config='--psm 7 digits').strip()
    print(f"CAPTCHA leído: {captcha_text}")

    # Ingresar el texto del CAPTCHA
    campo_captcha = driver.find_element(By.ID, "txtCaptcha")
    campo_captcha.clear()
    campo_captcha.send_keys(captcha_text)

    # Hacer clic en Consultar
    boton = driver.find_element(By.ID, "btnConsultar")
    boton.click()

    # Esperar los resultados (ajusta el tiempo o espera a un elemento)
    time.sleep(5)

    # Puedes extraer el resultado aquí o guardar la página
    # result = driver.page_source

    # Vuelve si deseas procesar otro
    # driver.back()

driver.quit()
