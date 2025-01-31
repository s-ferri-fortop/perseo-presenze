import pandas as pd
import fnmatch
from xml.dom import minidom
import os
from src.libraries.DipendentiCloud import Dipendente, Movimento
import datetime
from pathlib import Path
from zipfile import ZipFile
from os.path import basename
import numpy as np


class Presenze:
    def __init__(self):
        self.dict_cod_giust = self.compute_codice_giust_dip()

    def compute_codice_giust_dip(self):
        # df_cod_giust_dip = pd.read_excel(r'C:\xampp\htdocs\perseo-dev\fastprototyping\toolkits\Zucchetti\data\codice_giustificativo_dipendenti.xlsx')
        df_cod_giust_dip = pd.read_excel(
            Path(os.path.dirname(os.path.realpath(__file__)), "data/codice_giustificativo_dipendenti.xlsx",
                 encoding="utf-8"))
        from_dip = df_cod_giust_dip.loc[:, "FROM_DIPENDENTI IN CLOUD"].dropna()
        to_inf = df_cod_giust_dip.loc[:, "TO_INFINITY"].dropna()
        to_infinity = to_inf.tolist()
        from_dipendente = from_dip.tolist()
        dict_cod_giust = dict(zip(from_dipendente, to_infinity))
        return dict_cod_giust

    def read_file(self, file, file_name, month_number, year):
        dipendenti = []
        movimenti = []
        if fnmatch.fnmatch(file_name, '*.xlsx'):
            df = pd.read_excel(file)

            row_index, c_totale = np.where(df == "Totale")  # the length of days from the excel file, in this way we can remove the problem "day is out of range",
            len_col = c_totale[0]  # take the first element of c_totale array
            # df = pd.read_excel(str(path_month_data_folder) + "/" +file_name)

            for index, row in df.iterrows():
                value = row.iloc[0]
                if pd.isnull(value) or value == "Giustificativo":
                    continue
                elif value == "Matricola":
                    movimenti = []
                    name = row.iloc[3]
                    matricola = row.iloc[1]

                elif value == "Note":
                    if matricola != 999 and matricola != 998:  # Francesco e Claudia
                        dipendente = Dipendente(name, matricola, movimenti)
                        dipendenti.append(dipendente)
                elif value == "Protocolli di malattia":
                    continue
                else:
                    for i in range(1, len_col):
                        try:
                            date = datetime.date(year, month_number, i)
                        except ValueError:
                            break
                        nr_ore = row.iloc[i]
                        movimento = Movimento(value, date, nr_ore)
                        movimenti.append(movimento)
            return dipendenti

    def create_xml(self, dipendenti, month_number):
        # start with the creation of the xml file, the library that we are using is called xml.minidom
        # path_month_data_folder= Path(os.path.dirname(os.path.realpath(__file__), "data/", encoding="utf-8"))
        month_path = None
        if len(str(month_number)) == 1: # because int has not a length we convert it to a string
            month_path = Path(os.path.dirname(os.path.realpath(__file__)), "data/20220"+str(month_number), encoding="utf-8")
        elif len(str(month_number)) == 2:
            month_path = Path(os.path.dirname(os.path.realpath(__file__)), "data/2022" + str(month_number),
                             encoding="utf-8")
        if not os.path.isdir(month_path):
            os.mkdir(month_path)
        root = minidom.Document()
        xml = root.createElement('Fornitura')
        root.appendChild(xml)
        for i in range(0, len(dipendenti)):
            cod_dip_uff = dipendenti[i].codice_ufficiale_dipendente
            # if "matricola" of a dipendent does not exist or is equl to "nan", than do not write the content of that dipendent in the xml file
            if str(cod_dip_uff).endswith("nan"):
                print("MISSING CODE FOR: " + dipendenti[i].name)
                continue
            dipChild = root.createElement('Dipendente')
            dipChild.setAttribute('CodAziendaUfficiale', '000053')
            dipChild.setAttribute('CodDipendenteUfficiale', cod_dip_uff)
            xml.appendChild(dipChild)
            movimentiChild = root.createElement('Movimenti')
            movimentiChild.setAttribute('GenerazioneAutomaticaDaTeorico', 'N')
            dipChild.appendChild(movimentiChild)

            # create the calendar object taken from the Dipendente class
            calendar = dipendenti[i].calendar
            count = len(calendar)

            # make a loop in order to control for all key/value(day/movimenti) pairs of calendar
            for day, movimenti in calendar.items():
                # control if movimenti contains just one voice(in our example the element always present is "Presenza effetiva")
                # if it has just one voice we will write it in the xml file
                if len(movimenti.keys()) == 2:
                    gius = list(movimenti.keys())[1]
                    num_ore = list(movimenti.values())[0]
                    value = self.dict_cod_giust[gius]
                    code_name = root.createTextNode(value)
                    CodGiustificativoUfficiale = root.createElement('CodGiustificativoUfficiale')
                    CodGiustificativoUfficiale.appendChild(code_name)

                    movimentoChild = root.createElement('Movimento')
                    date = root.createElement('Data')
                    str_number_date = str(day)
                    number_date = root.createTextNode(str_number_date)
                    movimentoChild.appendChild(date)
                    number_date = root.createTextNode(str_number_date)
                    date.appendChild(number_date)

                    str_number_hours = str(num_ore)
                    number_minutes = None
                    if str_number_hours.endswith(".5"):
                        str_number_hours = str_number_hours.replace(".5", "")
                        str_number_minutes = "30"
                        number_minutes = root.createTextNode(str_number_minutes)
                    elif str_number_hours.endswith(".0"):
                        str_number_hours = str_number_hours.replace(".0", "")
                    number_hours = root.createTextNode(str_number_hours)

                    movimentoChild = root.createElement('Movimento')
                    movimentiChild.appendChild(movimentoChild)
                    movimentoChild.appendChild(CodGiustificativoUfficiale)
                    movimentoChild.appendChild(date)
                    date.appendChild(number_date)

                    num_ore = root.createElement('NumOre')
                    num_ore.appendChild(number_hours)
                    movimentoChild.appendChild(num_ore)
                    if number_minutes is not None:
                        num_min = root.createElement('NumMinuti')
                        num_min.appendChild(number_minutes)
                        movimentoChild.appendChild(num_min)


                else:
                    # if movimenti contain more than one element(value) we need to remove the "Presenza effettiva" voice
                    # from movimenti dict and write the next voices in xml file
                    if "Presenza effettiva" in movimenti:
                        in_office = movimenti["Presenza effettiva"]
                        # movimenti.pop("Presenza effettiva")
                    if "Presenza ordinaria prevista" in movimenti:
                        teorico = movimenti["Presenza ordinaria prevista"]
                        movimenti.pop("Presenza ordinaria prevista")
                    # if "Presenza effettiva" in movimenti:
                    #     gius = movimenti.pop("Presenza effettiva")
                    hours_sum = 0
                    for giust, num_ore in movimenti.items():
                        if giust != "Presenza effettiva":
                            hours_sum = hours_sum + float(num_ore)
                    if hours_sum == teorico:
                        try:
                            movimenti.pop("Presenza effettiva")
                        except Exception as e:
                            print(e)
                    else:
                        try:
                            difference = teorico - hours_sum
                            movimenti["Presenza effettiva"] = difference
                        except Exception as e:
                            print(e)

                    for giust, num_ore in movimenti.items():
                        if giust in self.dict_cod_giust.keys():
                            value = self.dict_cod_giust[giust]
                            code_name = root.createTextNode(value)
                            CodGiustificativoUfficiale = root.createElement('CodGiustificativoUfficiale')
                            CodGiustificativoUfficiale.appendChild(code_name)
                            movimentoChild = root.createElement('Movimento')
                            date = root.createElement('Data')
                            str_number_date = str(day)
                            number_date = root.createTextNode(str_number_date)
                            movimentoChild.appendChild(date)
                            number_date = root.createTextNode(str_number_date)
                            date.appendChild(number_date)

                            str_number_hours = str(num_ore)
                            number_minutes = None
                            if str_number_hours.endswith(".5"):
                                str_number_hours = str_number_hours.replace(".5", "")
                                str_number_minutes = "30"
                                number_minutes = root.createTextNode(str_number_minutes)
                            str_number_hours = str(int(num_ore))
                            number_hours = root.createTextNode(str_number_hours)

                            movimentoChild = root.createElement('Movimento')
                            movimentiChild.appendChild(movimentoChild)
                            movimentoChild.appendChild(CodGiustificativoUfficiale)
                            movimentoChild.appendChild(date)

                            date.appendChild(number_date)
                            num_ore = root.createElement('NumOre')
                            num_ore.appendChild(number_hours)
                            movimentoChild.appendChild(num_ore)
                            if number_minutes is not None:
                                num_min = root.createElement('NumMinuti')
                                num_min.appendChild(number_minutes)
                                movimentoChild.appendChild(num_min)
                    else:
                        continue

        xml_str = root.toprettyxml(indent="\t")
        file_path_to_save = "TO_"+str(month_number) + ".xml"
        with open(file_path_to_save, "w") as f:
            f.write(xml_str)
        return file_path_to_save


    ## il mio metodo
    def create_zip(self, month_number):
        output = str(Path(os.path.dirname(os.path.realpath(__file__)), "data/output", encoding="utf-8"))
        if len(str(month_number)) == 1:  # because int has not a length
            input = Path(os.path.dirname(os.path.realpath(__file__)), "data/20220"+str(month_number), encoding="utf-8")
        if len(str(month_number)) == 2:
            input = Path(os.path.dirname(os.path.realpath(__file__)), "data/2022" + str(month_number), encoding="utf-8")
        # create a ZipFile object
        try:
            with ZipFile(output+'/presenze.zip', 'w') as zipObj:
                for folderName, subfolders, filenames in os.walk(input):
                    for filename in filenames:
                        if filename.endswith('.xml'):
                            filePath = os.path.join(folderName, filename)
                            zipObj.write(filePath, basename(filePath))
                        else:
                            continue
            return zipObj
        except Exception as e:
            raise e
