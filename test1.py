from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from docx import Document
import time

# Configuración del driver
driver = webdriver.Chrome()  
driver.maximize_window()

# Navegar al formulario
driver.get("https://demoqa.com/automation-practice-form")

# Función para completar el formulario
def completar_formulario():
    resultados = {}

    try:
        # Llenar nombre y apellido
        driver.find_element(By.ID, "firstName").send_keys("Brenda")
        driver.find_element(By.ID, "lastName").send_keys("Tanoni")
        resultados["Nombre Completo"] = "Brenda Tanoni"

        # Llenar correo electrónico
        driver.find_element(By.ID, "userEmail").send_keys("brenda.tanoni@example.com")
        resultados["Email"] = "brenda.tanoni@example.com"

        # Seleccionar género
        driver.find_element(By.XPATH, "//label[text()='Female']").click()
        resultados["Género"] = "Female"

        # Llenar número móvil
        driver.find_element(By.ID, "userNumber").send_keys("1234567890")
        resultados["Teléfono"] = "1234567890"

        # Seleccionar fecha de nacimiento
        driver.find_element(By.ID, "dateOfBirthInput").click()
        Select(driver.find_element(By.CLASS_NAME, "react-datepicker__month-select")).select_by_visible_text("May")
        Select(driver.find_element(By.CLASS_NAME, "react-datepicker__year-select")).select_by_visible_text("1995")
        driver.find_element(By.XPATH, "//div[text()='15']").click()
        resultados["Fecha de Nacimiento"] = "15 May 1995"

        # Llenar asignaturas
        driver.find_element(By.ID, "subjectsInput").send_keys("Maths\n")
        resultados["Asignaturas"] = "Maths"

        # Seleccionar hobbies
        driver.find_element(By.XPATH, "//label[text()='Sports']").click()
        resultados["Hobbies"] = "Sports"

        # Dirección
        driver.find_element(By.ID, "currentAddress").send_keys("San Fernando, Buenos Aires")
        resultados["Dirección"] = "San Fernando, Buenos Aires"

        # Exportar resultados
        exportar_a_word(resultados)
        print("Formulario completado y resultados exportados con éxito.")
    
    except Exception as e:
        print(f"Error durante la automatización: {e}")
    finally:
        driver.quit()

# Función para exportar resultados a un archivo Word
def exportar_a_word