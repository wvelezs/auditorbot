import os
import time
import json
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import pandas.io.excel._xlsxwriter

ol = os.environ.get('USERNAME')


class Modificarexcel():
    nuevoexcel = ""

    def leerexcel(self):
        self.nuevoexcel = pd.read_excel(
            'C:/Users/'+ol+'/Dropbox/INFORMES ALTO COSTO/SANITAS/infoavicena - copia.xlsx', sheet_name="Hoja1")

    def test_modificarexcel(self):
        # médico
        self.nuevoexcel[["col0001", "col0002"]] = self.nuevoexcel.SIGNOS.str.rsplit(
            ". Reg.", n=1, expand=True)  # genera frecuencia cardíaca
        self.nuevoexcel[["col0003", "MÉDICO"]] = self.nuevoexcel.col0001.str.rsplit(
            ".\n", n=1, expand=True)
        self.nuevoexcel.drop(
            columns=["col0001", "col0002", "col0003"], inplace=True)
        # tipo de historia:
        self.nuevoexcel.dropna(inplace=False)
        # crea dos columnas = de la columna SIGNOS haciendo split en el delimitador "telefono:" #corta la info de la columna SIGNOS, n=1 corta una sola vez, expand true permite que se generen las sig columnas

        self.nuevoexcel[["col001", "col002"]] = self.nuevoexcel.SIGNOS.str.split(
            "S.A.S\n", n=1, expand=True)
        self.nuevoexcel[["TIPO DE HISTORIA", "col003"]] = self.nuevoexcel.col002.str.split(
            ".", n=1, expand=True)  # TELEFONO crea la columna con la info limpia, cuatro es una columna residuo
        self.nuevoexcel.drop(
            columns=["col001", "col002", "col003"], inplace=True)

        self.nuevoexcel[["colxxx1", "colxxx2"]] = self.nuevoexcel.SIGNOS.str.split(
            "DIAGNÓSTICO\n\n", n=1, expand=True)
        self.nuevoexcel[["DIAGNÓSTICO", "colxxx3"]] = self.nuevoexcel.colxxx2.str.split(
            "\n\nPLAN", n=1, expand=True)  # TELEFONO crea la columna con la info limpia, cuatro es una columna residuo
        # elimina las columnas basura
        self.nuevoexcel.drop(
            columns=["colxxx1", "colxxx2", "colxxx3"], inplace=True)

        # fecha:
        self.nuevoexcel.dropna(inplace=False)
        # crea dos columnas = de la columna SIGNOS haciendo rsplit en el delimitador "responsable:" #corta la info de la columna SIGNOS, n=1 corta una sola vez, expand true permite que se generen las sig columnas
        self.nuevoexcel[["col01", "col02"]] = self.nuevoexcel.SIGNOS.str.rsplit(
            ". Responsable:", n=1, expand=True)
        self.nuevoexcel[["col03", "FECHA"]] = self.nuevoexcel.col01.str.rsplit(
            ". ", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col01", "col02", "col03"], inplace=True)
        # Telefono:
        self.nuevoexcel.dropna(inplace=False)

        # crea dos columnas = de la columna SIGNOS haciendo split en el delimitador "telefono:" #corta la info de la columna SIGNOS, n=1 corta una sola vez, expand true permite que se generen las sig columnas
        self.nuevoexcel[["uno", "dos"]] = self.nuevoexcel.SIGNOS.str.split(
            "Telefono: ", n=1, expand=True)
        # TELEFONO crea la columna con la info limpia, cuatro es una columna residuo
        self.nuevoexcel[["TELEFONO", "cuatro"]
                        ] = self.nuevoexcel.dos.str.split(".", n=1, expand=True)
        # elimina las columnas basura
        self.nuevoexcel.drop(columns=["uno", "dos", "cuatro"], inplace=True)

        self.nuevoexcel[["5", "seis"]] = self.nuevoexcel.SIGNOS.str.split(
            "Estado general: ", n=1, expand=True)  # genera estado general
        self.nuevoexcel[["ESTADO GENERAL", "8"]] = self.nuevoexcel.seis.str.split(
            "\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["5", "seis", "8"], inplace=True)

        self.nuevoexcel[["nueve", "diez"]] = self.nuevoexcel.SIGNOS.str.split(
            "Frecuencia Cardíaca: ", n=1, expand=True)  # genera frecuencia cardíaca
        self.nuevoexcel[["FRECUENCIA CARDÍACA  (Latidos/min)", "once"]
                        ] = self.nuevoexcel.diez.str.split(" Latidos/min", n=1, expand=True)
        self.nuevoexcel.drop(columns=["nueve", "diez", "once"], inplace=True)

        self.nuevoexcel[["doce", "trece"]] = self.nuevoexcel.SIGNOS.str.split(
            "Frecuencia Respiratoria: ", n=1, expand=True)  # Frecuencia Respiratoria:
        self.nuevoexcel[["FRECUENCIA RESPIRATORIA (Respiraciones/min)", "cato"]
                        ] = self.nuevoexcel.trece.str.split(" Respiraciones/min\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["doce", "trece", "cato"], inplace=True)

        self.nuevoexcel[["quince", "dieciseis"]] = self.nuevoexcel.SIGNOS.str.split(
            "Tensión Arterial Sistólica: ", n=1, expand=True)  # genera Tensión Arterial Sistólica:
        self.nuevoexcel[["TENSIÓN ARTERIAL SISTÓLICA (mmHg)", "diesy7"]] = self.nuevoexcel.dieciseis.str.split(
            " mmHg\n", n=1, expand=True)
        self.nuevoexcel.drop(
            columns=["quince", "dieciseis", "diesy7"], inplace=True)

        self.nuevoexcel[["col18", "col19"]] = self.nuevoexcel.SIGNOS.str.split(
            "Tensión Arterial Diastólica: ", n=1, expand=True)  # genera Tensión Arterial Diastólica:
        self.nuevoexcel[["TENSIÓN ARTERIAL DIASTÓLICA (mmHg)", "col20"]] = self.nuevoexcel.col19.str.split(
            " mmHg\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col18", "col19", "col20"], inplace=True)

        self.nuevoexcel[["col21", "col22"]] = self.nuevoexcel.SIGNOS.str.split(
            "Tensión Arterial Media: ", n=1, expand=True)  # genera Tensión Arterial media:
        self.nuevoexcel[["TENSIÓN ARTERIAL MEDIA (mmHg)", "col23"]] = self.nuevoexcel.col22.str.split(
            " mmHg\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col21", "col22", "col23"], inplace=True)

        self.nuevoexcel[["col27", "col28"]] = self.nuevoexcel.SIGNOS.str.split(
            "Temperatura: ", n=1, expand=True)  # genera Temperatura:
        self.nuevoexcel[["TEMPERATURA (ºC)", "col29"]] = self.nuevoexcel.col28.str.split(
            " ºC\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col27", "col28", "col29"], inplace=True)

        self.nuevoexcel[["col30", "col31"]] = self.nuevoexcel.SIGNOS.str.split(
            "Peso: ", n=1, expand=True)  # genera Peso:
        self.nuevoexcel[["PESO ( Kg)", "col32"]] = self.nuevoexcel.col31.str.split(
            " Kg\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col30", "col31", "col32"], inplace=True)

        self.nuevoexcel[["col33", "col34"]] = self.nuevoexcel.SIGNOS.str.split(
            "Talla: ", n=1, expand=True)  # genera Talla:
        self.nuevoexcel[["TALLA", "col35"]] = self.nuevoexcel.col34.str.split(
            "\n", n=1, expand=True)

        self.nuevoexcel.drop(
            columns=["col33", "col34", "col35"], inplace=True)

        self.nuevoexcel[["col36", "col37"]] = self.nuevoexcel.SIGNOS.str.split(
            "Índice de Masa Corporal: ", n=1, expand=True)  # genera Índice de Masa Corporal:
        self.nuevoexcel[["IMC", "col38"]] = self.nuevoexcel.col37.str.split(
            "-", n=1, expand=True)

        self.nuevoexcel.drop(
            columns=["col36", "col37", "col38"], inplace=True)

        self.nuevoexcel[["col036", "col037"]] = self.nuevoexcel.SIGNOS.str.split(
            "\) -", n=1, expand=True)  # genera CLASIFICACIÓN PESO
        self.nuevoexcel[["CLASIFICACIÓN PESO", "col035"]
                        ] = self.nuevoexcel.col037.str.split("\n", n=1, expand=True)

        self.nuevoexcel.drop(
            columns=["col036", "col037", "col035"], inplace=True)

        self.nuevoexcel[["col39", "col40"]] = self.nuevoexcel.SIGNOS.str.split(
            "Circunferencia de la cintura: ", n=1, expand=True)  # genera Circunferencia de la cintura:
        self.nuevoexcel[["CIRCUNFERENCIA DE LA CINTURA", "col41"]
                        ] = self.nuevoexcel.col40.str.split("\n", n=1, expand=True)

        self.nuevoexcel.drop(
            columns=["col39", "col40", "col41"], inplace=True)

        self.nuevoexcel[["col42", "col43"]] = self.nuevoexcel.SIGNOS.str.split(
            "Superficie corporal: ", n=1, expand=True)  # genera Superficie corporal:
        self.nuevoexcel[["SUPERFICIE CORPORAL", "col44"]
                        ] = self.nuevoexcel.col43.str.split("\n", n=1, expand=True)
        self.nuevoexcel.drop(columns=["col42", "col43", "col44"], inplace=True)
        self.nuevoexcel.to_excel(
            'C:/Users/'+ol+'/Dropbox/INFORMES ALTO COSTO/SANITAS/infoavicena - copia.xlsx', sheet_name='Hoja1', index=False)


obje = Modificarexcel()
obje.leerexcel()
obje.test_modificarexcel()
