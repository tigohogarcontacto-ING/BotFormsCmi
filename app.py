from flask import Flask, render_template, request, redirect, url_for
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os, json, time

app = Flask(__name__)

# --- Configuración Google Sheets ---
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

creds_json = os.environ.get("GOOGLE_CREDENTIALS")
if creds_json:
    creds_dict = json.loads(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
else:
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)

client = gspread.authorize(creds)
SPREADSHEET_ID = "1-RXO6IkgxMHBnRlycH7dExeozQwFCb4vxDNL-i8dAEo"
worksheet = client.worksheet("CMIA")


def enviar_fila(fila):
    """Toma una fila de Google Sheets y la envía al Microsoft Forms"""

    # instala automáticamente chromedriver
    chromedriver_autoinstaller.install()

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://forms.office.com/r/JMfLJRFdJe")
    time.sleep(3)

    # Fecha fija
    fecha = driver.find_element(By.ID, "DatePicker0-label")
    fecha.send_keys("09/09/2025")

    # Campos numéricos
    campos_numericos = driver.find_elements(By.XPATH, "//input[@placeholder='Escriba un número entero']")
    campos_numericos[0].send_keys("1113860034")  # Cédula agente fija

    # Radios fijos
    driver.find_element(By.XPATH, "//input[@type='radio' and @value='BECALL']").click()
    driver.find_element(By.XPATH, "//input[@type='radio' and @value='Salida']").click()
    driver.find_element(By.XPATH, "//input[@type='radio' and @value='Vtex']").click()
    driver.find_element(By.XPATH, "//input[@type='radio' and @value='Especializado']").click()
    driver.find_element(By.XPATH, "//input[@type='radio' and @value='ANDINA']").click()

    # Texto libre desde Sheets
    driver.find_element(
        By.XPATH,
        "//input[@aria-labelledby='QuestionId_rd11883bb95014c7fa475fe5b749b62a3 QuestionInfo_rd11883bb95014c7fa475fe5b749b62a3']"
    ).send_keys(fila["DEPARTAMENTO"])

    driver.find_element(
        By.XPATH,
        "//input[@aria-labelledby='QuestionId_r78dcba4ea8df4571a7541e9389120efe QuestionInfo_r78dcba4ea8df4571a7541e9389120efe']"
    ).send_keys(fila["CIUDAD"])

    driver.find_element(By.XPATH, "//input[@type='radio' and @value='PORTABILIDAD']").click()

    # Número y cédula desde Sheets
    campos_numericos[1].send_keys(str(fila["NUMERO"]))
    campos_numericos[2].send_keys(str(fila["CEDULA"]))

    # Botón Enviar
    driver.find_element(By.XPATH, "//button[@data-automation-id='submitButton']").click()
    time.sleep(2)
    driver.quit()


@app.route("/", methods=["GET"])
def index():
    datos = worksheet.get_all_records()
    return render_template("index.html", datos=datos)


@app.route("/enviar", methods=["POST"])
def enviar():
    datos = worksheet.get_all_records()
    for fila in datos:
        enviar_fila(fila)
    return render_template("resultado.html", total=len(datos))


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
