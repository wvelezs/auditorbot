from tfidf import TfIdf
import os
import sys
import time
import json
import numpy as np
import pandas as pd
import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import Select
import pandas.io.excel._xlsxwriter

sys.path.insert(1, 'C:/Users/Soporte/Documents/chatbotpyp/MyScripts')
ol = os.environ.get('USERNAME')


class TestUntitled():
    cc = []
    tip = []
    tipodoc = {"CC": "Cédula de ciudadanía", "CE": "Cédula de extranjería", "TI": "Tarjeta de identidad",
               "PS": "Pasaporte", "PE": "Permiso Especial de Permanencia"}

    def setup_method(self, method):
        self.driver = webdriver.Chrome(
            'C:/Users/'+ol+'/Dropbox/chromedriver_win32/chromedriver.exe')
        self.vars = {}

    def leerexcel(self):
        df = pd.read_excel(
            'usuarios.xlsx', sheet_name="Hoja1")

        for i in df.index:
            self.cc.append(str(df['cedulas'][i]))
            self.tip.append(str(df['tipo'][i]))

    def teardown_method(self, method):

        self.driver.quit()

    def wait_for_window(self, timeout=2):
        time.sleep(round(timeout / 1000))
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        if len(wh_now) > len(wh_then):
            return set(wh_now).difference(set(wh_then)).pop()

    def test_untitled(self):
        self.driver.get("https://avicena.colsanitas.com/His/login.seam")
        self.driver.set_window_size(1064, 816)
        self.driver.find_element(
            By.CSS_SELECTOR, "#pnlPrincipalLogin > .rowForm:nth-child(1)").click()
        self.driver.find_element(
            By.ID, "ctlFormLogin:idUserNameLogin").send_keys("wijvelez")
        self.driver.find_element(
            By.ID, "ctlFormLogin:idPwdLogin").send_keys("Suarez2019")
        self.driver.find_element(By.ID, "ctlFormLogin:botonValidar").click()
        time.sleep(3)
        self.driver.find_element(By.ID, "formIngreso:btnIngresar").click()
        time.sleep(3)
        self.obtenerinfo()

    def obtenerinfo(self):
        datos = [{"03": [{"tipd": [], "signos":[], "cc":[], "nombre":[], }], "04":[{"tipd": [], "signos":[], "cc":[], "nombre":[], }], "05":[
            {"tipd": [], "signos":[], "cc":[], "nombre":[], }], "06":[{"tipd": [], "signos":[], "cc":[], "nombre":[], }]}]
        fechas = ["03", "04", "05", "06"]
        print(range(len(self.cc)))
        for i in range(len(self.cc)):
            try:
                self.driver.get(
                    "https://avicena.colsanitas.com/His/home.seam?cid=28449")
                self.driver.find_element(
                    By.CSS_SELECTOR, ".tablaGenericaPrincipal td:nth-child(2)").click()
                time.sleep(2)
                self.driver.find_element(
                    By.ID, "formMenu:ctlItemConsultarHistoria:anchor").click()
                time.sleep(2)
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlTipoIdentificacion").click()
                dropdown = self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlTipoIdentificacion")
                dropdown.find_element(
                    By.XPATH, "//option[. = '" + self.tipodoc[self.tip[i]] + "']").click()
                time.sleep(2)
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlTipoIdentificacion").click()
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlNumeroDocumento").click()
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlNumeroDocumento").send_keys(self.cc[i])
                time.sleep(2.5)
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:buscar1").click()
                time.sleep(3)
                self.driver.find_element(
                    By.ID, "formPrerrequisitos:ctlFolios:0:ctlConsultaHC").click()
                time.sleep(3)
                self.driver.find_element(
                    By.ID, "formModalFMFormulasNoCatalogo:justificacionCombo").click()
                dropdown = self.driver.find_element(
                    By.ID, "formModalFMFormulasNoCatalogo:justificacionCombo")
                time.sleep(3)
                dropdown.find_element(
                    By.XPATH, "//option[. = 'Auditoría de historia clínica']").click()
                self.driver.find_element(
                    By.ID, "formModalFMFormulasNoCatalogo:justificacionCombo").click()
                time.sleep(3)
                self.vars["window_handles"] = self.driver.window_handles
                self.driver.find_element(
                    By.ID, "formModalFMFormulasNoCatalogo:ctlAceptar").click()
                self.vars["win815"] = self.wait_for_window(4000)
                self.vars["root"] = self.driver.current_window_handle
                self.driver.switch_to.window(self.vars["win815"])
                nomb = self.driver.find_element(
                    By.XPATH, "//div[@id=\'divDatosPacientes\']/div/table/tbody/tr/td[4]/span").text
                print(nomb)
                time.sleep(2.5)
                self.driver.find_element(
                    By.LINK_TEXT, "Últimos eventos").click()
                time.sleep(3)

                for x in range(5):
                    time.sleep(3)
                    fecha = self.driver.find_element(
                        By.ID, "formHC:rptOtros:"+str(x)+":j_id148").text
                    dato = [fecha.split("/")[1]]
                    print(dato)

                    resultado = self.calcularfrecuencia(fechas, dato)

                    print(resultado)
                    if(resultado == True):
                        time.sleep(3)
                        self.driver.find_element(
                            By.ID, "formHC:rptOtros:"+str(x)+":j_id148").click()
                        time.sleep(3)
                        pruebaaa = str(self.driver.find_element_by_id(
                            "formHC:ctlResumen").get_attribute('value'))
                        time.sleep(3)
                        datos[0][dato[0]][0]["tipd"].append(str(self.tip[i]))
                        datos[0][dato[0]][0]["cc"].append(str(self.cc[i]))
                        datos[0][dato[0]][0]["nombre"].append(str(nomb))
                        datos[0][dato[0]][0]["signos"].append(pruebaaa)
                        time.sleep(3)
                self.driver.close()
                self.driver.switch_to.window(self.vars["root"])
            except:
                self.driver.get(
                    "https://avicena.colsanitas.com/His/home.seam?cid=28449")
                print('hubo error '+str(i))

                if i == 0:
                    i = 0
                    pass
                else:
                    i -= 1
                    pass

                print('hubo error '+str(i)+' '+str(self.cc[i]))
            pass
        self.saveExcel(datos, fechas)
        print('ya terminó')

    def calcularfrecuencia(self, texto, palabra=[]):

        table = TfIdf()
        table.add_document("informacion", texto)
        resultado = table.similarities(palabra)[0][1]
        if resultado > 0.0:
            return True
        return False

    def saveExcel(self, datos, fechas):
        for fecha in fechas:
            df = pd.DataFrame(
                {'TIPO':  datos[0][fecha][0]["tipd"], 'CC':  datos[0][fecha][0]["cc"], 'NOMBRE':  datos[0][fecha][0]["nombre"], 'SIGNOS': datos[0][fecha][0]["signos"]})
            df.to_excel(fecha+'.xlsx',
                        sheet_name='Hoja1', index=False)


obje = TestUntitled()
obje.setup_method("get")
obje.leerexcel()
obje.test_untitled()
obje.teardown_method("get")
