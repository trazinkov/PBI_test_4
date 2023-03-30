# -*- coding: utf-8 -*-
# convert ITWO codes into 7 levels of text parameters
# get object with 3 parameters from ITWO (csv file)
# and return 3 + 7 parameters (for every sorting level)

'''add lib'''
import sys
sys.path += [
    r"Q:\DESTUALB\15-BIM5D\10-B5D\22-Hochbau\60-ENTWICKLUNG\Dynamo\Templates\Python\modules_Anatoly"]  # noqa

file = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 10 - Netto and Brutto\Excel\ITWO connection\ITWO_Code_ID_Value.csv"  # noqa


class ITWO_record():
    """ Object represent one string in csv file. Code, ID and Vlaue
    Code is analysing and return set of 6 parameters for every ITWO_record.
    Itis not revit element. Seweral ITWO_elements coould have the same 
    element ID but different construction codes and parameters."""

    def __init__(self, construction_code, Id, value):
        self.construction_code = construction_code
        self.Id = Id
        self.value = value

        self.level_1 = self.construction_code.split(".")[0]
        self.level_2 = self.construction_code.split(".")[1]
        self.level_3 = self.construction_code.split(".")[2]
        self.level_4 = self.construction_code.split(".")[3]
        self.level_5 = self.construction_code.split(".")[4]
        self.level_6 = self.construction_code.split(".")[5]

        self.level_1_text = ""
        self.level_2_text = ""
        self.level_3_text = ""
        self.level_4_text = ""
        self.level_5_text = ""
        self.level_6_text = ""

        if self.level_1 == "1":
            self.level_1_text = "Plausibilitätsmengen"

            if self.level_2 == "01":
                self.level_2_text = "Estrich"
                if self.level_3 == "01":
                    self.level_3_text = "Zementestrich"
                    if self.level_4 == "00":
                        self.level_4_text = "Zementestrich - Gesamt Verbund-,Trenn-,schwimmend-, Heizestrich"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Menge GESAMT"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Menge GESAMT"
                    elif self.level_4 == "05":
                        self.level_4_text = "Schwimmender Zementestrich"
                        if self.level_5 == "20":
                            self.level_5_text = "Netto Menge Trittschalldämmschicht"
                        elif self.level_5 == "30":
                            self.level_5_text = "Netto Menge TS + WD- Dämmschicht"
                    elif self.level_4 == "10":
                        self.level_4_text = "Einbauteile und sonstige Leistungen"
                        if self.level_5 == "40":
                            self.level_5_text = "Netto-Menge Belegreife"
                elif self.level_3 == "04":
                    self.level_3_text = "Hohlraumböden - gem. MLV 23.06.2021"
                    if self.level_4 == "00":
                        self.level_4_text = "Hohlraumböden - Gesamt Trocken-, Nassbauweise"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Mengen GESAMT"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Mengen GESAMT"
                    elif self.level_4 == "03":
                        self.level_4_text = "Hohlraumböden in Trockenbauweise"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto Menge HoBo Untergrundbehandlung Boden"
                        elif self.level_5 == "20":
                            self.level_5_text = "Brutto analog Revit Menge HoBo Untergrundbehandlung Umfang"
                    elif self.level_4 == "04":
                        self.level_4_text = "Hohlraumböden in Nassbauweise"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto Menge Hohlraumböden in Nassbauweise"

            elif self.level_2 == "03":
                self.level_2_text = "Boden- und Wandbeläge"
                if self.level_3 == "03":
                    self.level_3_text = "Boden- und Wandbeläge"
                    if self.level_4 == "03":
                        self.level_4_text = "Wandfliesen"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto Menge Wandfliesen"

            elif self.level_2 == "05":
                self.level_2_text = "Trockenbau / Verkleidungen"

                if self.level_3 == "01":
                    self.level_3_text = "Gipskartonständerwände"

                    if self.level_4 == "01":
                        self.level_4_text = "Metallständerwände ESW, DSW, DS, INS"
                        if self.level_5 == "00":
                            self.level_5_text = "Metallständerwände - Beidseitig Beplankt - Gesamt ESW+DSW+INS+DS"
                            if self.level_6 == "0090":
                                self.level_6_text = "Netto Menge Metallständerwände - beidseitige Beplankung - Gesamt ESW+DSW+INS+DS"
                            elif self.level_6 == "0100":
                                self.level_6_text = "VOB Menge Metallständerwände - beidseitige Beplankung - Gesamt ESW+DSW+INS+DS"
                        elif self.level_5 == "01":
                            self.level_5_text = "Einfachständerwände"
                            if self.level_6 == "0011":
                                self.level_6_text = "VOB Menge Metallständerwände - ESW"
                            elif self.level_6 == "0015":
                                self.level_6_text = "Netto Menge Metallständerwände - ESW"

                    elif self.level_4 == "02":
                        self.level_4_text = "Schachtwände und Vorsatzschalen"
                        if self.level_5 == "00":
                            self.level_5_text = "Metallständerwände - Einseitig Beplankt - Gesamt SCHA + VS"
                            if self.level_6 == "0110":
                                self.level_6_text = "Netto Menge Metallständerwände - Einseitige Beplankung - Gesamt SCHA+VS"
                            elif self.level_6 == "0125":
                                self.level_6_text = "VOB Menge Metallständerwände - Einseitige Beplankung - Gesamt SCHA+VS"
                        elif self.level_5 == "01":
                            self.level_5_text = "Schachtwände"
                            if self.level_6 == "0010":
                                self.level_6_text = "Netto Menge Schachtwände"
                            elif self.level_6 == "0020":
                                self.level_6_text = "VOB Menge Schachtwände"
                        elif self.level_5 == "02":
                            self.level_5_text = "Vorsatzschalen"
                            if self.level_6 == "0072":
                                self.level_6_text = "Netto Menge Vorsatzschalen freistehend raumhoch"
                            elif self.level_6 == "0082":
                                self.level_6_text = "VOB Menge Vorsatzschalen freistehend raumhoch"
                            elif self.level_6 == "0110":
                                self.level_6_text = "Netto-Menge Vorsatzschale GESAMT"

                    elif self.level_4 == "06":
                        self.level_4_text = "Mehrpreise/ Einbauten Gispkartonständerwände F0-F90"
                        if self.level_5 == "01":
                            self.level_5_text = "Gleitende Deckennschluss Plausi"
                            if self.level_6 == "0010":
                                self.level_6_text = "VOB Gleitender Deckenanschluss F0-F90 - Plausi"
                            elif self.level_6 == "0012":
                                self.level_6_text = "Netto Gleitender Deckenanschluss F0-F90 - Plausi"

                elif self.level_3 == "03":
                    self.level_3_text = "Gipskartondecken"
                    if self.level_4 == "00":
                        self.level_4_text = "Gipskartondecken Gesamt"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Menge GESAMT"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Menge GESAMT"
                    elif self.level_4 == "01":
                        self.level_4_text = "Gipskartonplattendecken"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto Menge Gipskartonplattendecken"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB Menge Gipskartondeckenplatten"

                elif self.level_3 == "05":
                    self.level_3_text = "Metalldecken"
                    if self.level_4 == "00":
                        self.level_4_text = "Metalldecken - Gesamt Kassetten-, Paneel-, Lamellen- und Streckmetall"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Menge GESAMT"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Menge GESAMT"
                    elif self.level_4 == "01":
                        self.level_4_text = "Kassettendecken"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Menge Kassettendecken"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Menge Kassettendecken"
                    elif self.level_4 == "02":
                        self.level_4_text = "Paneeledecken"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto-Menge Paneeledecken"
                        elif self.level_5 == "20":
                            self.level_5_text = "VOB-Menge Paneeledecken"

            elif self.level_2 == "06":
                self.level_2_text = "Einbauwände"
                if self.level_3 == "01":
                    self.level_3_text = "Systemtrennwände"
                    if self.level_4 == "01":
                        self.level_4_text = "Systemtrenn-Wände"
                        if self.level_5 == "10":
                            self.level_5_text = "Menge Ganzglaselemente"
                        elif self.level_5 == "27":
                            self.level_5_text = "Gesamtmenge Systemtüren Alu"
                elif self.level_3 == "02":
                    self.level_3_text = "WC-Trennwände"
                    if self.level_4 == "01":
                        self.level_4_text = "WC-Trennwände"
                        if self.level_5 == "30":
                            self.level_5_text = "Gesamtmenge WC-Türen"
                elif self.level_3 == "03":
                    self.level_3_text = "Mobile Schiebewände"
                    if self.level_4 == "01":
                        self.level_4_text = "Holz-Element Wände"
                        if self.level_5 == "10":
                            self.level_5_text = "Netto Menge Holz-Element Wände"

            elif self.level_2 == "07":
                self.level_2_text = "Türen"
                if self.level_3 == "01":
                    self.level_3_text = "Holztüren"
                    if self.level_4 == "01":
                        self.level_4_text = "Holztüren"
                        if self.level_5 == "10":
                            self.level_5_text = "Gesamtmenge Holztüren"
                elif self.level_3 == "02":
                    self.level_3_text = "Stahlblechtüren"
                    if self.level_4 == "01":
                        self.level_4_text = "Stahlblechtüren"
                        if self.level_5 == "10":
                            self.level_5_text = "Gesamtmenge Stahlblechtüren"
                elif self.level_3 == "06":
                    self.level_3_text = "Alurahmentüren"
                    if self.level_4 == "01":
                        self.level_4_text = "Alurahmentüren"
                        if self.level_5 == "10":
                            self.level_5_text = "Gesamtmenge Alurahmentüren"

            elif self.level_2 == "09":
                self.level_2_text = "Malerarbeiten - gem. MLV 21.07.2021"
                if self.level_3 == "01":
                    self.level_3_text = "Anstriche innen"
                    if self.level_4 == "03":
                        self.level_4_text = "Tapezierarbeiten als Zwischenschicht"
                        if self.level_5 == "00":
                            self.level_5_text = "Tapezierarbeiten als Zwischenschicht GESAMT"
                            if self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Wände+ Stützen GESAMT"
                        elif self.level_5 == "02":
                            self.level_5_text = "Wände"
                            if self.level_6 == "0010":
                                self.level_6_text = "VOB-Menge Tapete auf Wände GESAMT"
                            elif self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Tapete auf Wände GESAMT"
                elif self.level_3 == "02":
                    self.level_3_text = "Spachtelarbeiten"
                    if self.level_4 == "01":
                        self.level_4_text = "Spachtelarbeiten als Unterschicht"
                        if self.level_5 == "00":
                            self.level_5_text = "Spachtelarbeiten GESAMT"
                            if self.level_6 == "0010":
                                self.level_6_text = "Netto-Menge Wände+Stützen GESAMT"
                            elif self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Decken GESAMT"
                            elif self.level_6 == "0030":
                                self.level_6_text = "Netto-Menge Unter+Überzüge GESAMT"
                            elif self.level_6 == "0040":
                                self.level_6_text = "Netto-Mengen Treppen/ Podeste GESAMT"
                        elif self.level_5 == "01":
                            self.level_5_text = "Spachtelarbeiten Wände"
                            if self.level_6 == "0010":
                                self.level_6_text = "VOB-Menge Spachtelarbeiten auf Ortbetonwänden"
                            elif self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Spachtelarbeiten auf Ortbetonwänden"
                        elif self.level_5 == "02":
                            self.level_5_text = "Spachtelarbeiten Decke"
                            if self.level_6 == "0010":
                                self.level_6_text = "VOB-Menge Spachtelarbeiten auf Decken GESAMT"
                            elif self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Spachtelarbeiten auf Decken GESAMT"
                        elif self.level_5 == "04":
                            self.level_5_text = "Spachtelarbeiten Unterzüge (Wangen) --> im LV unter Decken"
                            if self.level_6 == "0010":
                                self.level_6_text = "VOB-Menge Spachtelarbeiten auf Unterzug Wangen GESAMT"
                            elif self.level_6 == "0020":
                                self.level_6_text = "Netto-Menge Spachtelarbeiten auf Unterzug Wangen GESAMT"
                        elif self.level_5 == "06":
                            self.level_5_text = "Spachtelarbeiten Treppe"
                            if self.level_6 == "0030":
                                self.level_6_text = "VOB-Menge Spachtelarbeiten auf Treppen- und Podest-Untersichten GESAMT"
                            elif self.level_6 == "0040":
                                self.level_6_text = "Netto-Menge Spachtelarbeiten auf Treppen- und Podest-Untersichten GESAMT"

            elif self.level_2 == "10":
                self.level_2_text = "Sonderelemente"
                if self.level_3 == "01":
                    self.level_3_text = "Boden"
                    if self.level_4 == "03":
                        self.level_4_text = "Linie"
                        if self.level_5 == "10":
                            self.level_5_text = "Gesamtlänge Boden"

            elif self.level_2 == "11":
                self.level_2_text = "Fenster"
                if self.level_3 == "01":
                    self.level_3_text = "Leibung"
                    if self.level_4 == "03":
                        self.level_4_text = "Deckschicht"
                        if self.level_5 == "90":
                            self.level_5_text = "Anzahl Fenster Sonstiges"
                elif self.level_3 == "02":
                    self.level_3_text = "Fensterbank Material"
                    if self.level_4 == "01":
                        self.level_4_text = "Fensterbank Material"
                        if self.level_5 == "60":
                            self.level_5_text = "Gesamtbreite Fensterbank Sonstiges"


construction_codes = []
element_ids = []
values = []
strings = []

# collect info from csv file:
with open(file, 'r') as f:
    for x in f:
        construction_codes.append(str(x.split(",")[0]).split("\n")[0])
        element_ids.append(str(x.split(",")[1]).split("\n")[0])
        values.append(str(x.split(",")[2]).split("\n")[0])
        strings.append(str(x))

# collect ITWO records:
ITWO_records = [ITWO_record(construction_code, Id, value)
                for construction_code, Id, value
                in zip(construction_codes[1:], element_ids[1:], values[1:])]

# for every element ID return set of ITWO parameters, value etc:
data_list = []
for el in ITWO_records:
    element_list = []
    element_list.append(el.construction_code)
    element_list.append(el.Id)
    element_list.append(el.value)
    element_list.append(el.level_1_text)
    element_list.append(el.level_2_text)
    element_list.append(el.level_3_text)
    element_list.append(el.level_4_text)
    element_list.append(el.level_5_text)
    element_list.append(el.level_6_text)
    data_list.append(element_list)

OUT = data_list
# OUT = "Test"
