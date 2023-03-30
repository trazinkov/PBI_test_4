""" Analyze Revit document
to create data list with ID elements and Strabag parameters 
from Tables schedule """

import clr

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit import DB

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager


'''add lib'''
import sys
sys.path += [
    r"Q:\DESTUALB\15-BIM5D\10-B5D\22-Hochbau\60-ENTWICKLUNG\Dynamo\Templates\Python"]  # noqa

from modules_Anatoly.mod_Revit import type_parameter_value_for_instance
from modules_Anatoly.mod_Python import flatten


class PBI_element():
    """ Connection element btween Revit and Power BI (PBI).
    It collect properties from element, type element and starbag settings
    for elements, wich not used in the Revit (
    we have no such parameters in the Revit,
    like 'Bodenbeläge', 'Wandfliesen' etc).
    Connection according the excel summary scedule and
    Funcions from excel.rb file"""

    def __init__(self, doc, revit_element):
        self.doc = doc
        self.revit_element = revit_element
        self.category = self.revit_element.Category

    # create empty Strabag parameters
        self.Estriche_Bodensysteme_Lvl_1 = ""
        self.Estriche_Bodensysteme_Lvl_2 = ""

        self.Oberflächenschutz_Estrich_Beton_Lvl_1 = ""
        self.Oberflächenschutz_Estrich_Beton_Lvl_2 = ""

        self.Bodenbeläge_Lvl_1 = ""
        self.Bodenbeläge_Lvl_2 = ""

        self.Wandbeläge_Lvl_1 = ""
        self.Wandbeläge_Lvl_2 = ""

        self.Innenputz_Lvl_1 = ""
        self.Innenputz_Lvl_2 = ""

        self.Trockenbau_Verkleidungen_Wand_Lvl_1 = ""
        self.Trockenbau_Verkleidungen_Wand_Lvl_2 = ""

        self.Trockenbau_Verkleidungen_Decke_Lvl_1 = ""
        self.Trockenbau_Verkleidungen_Decke_Lvl_2 = ""

        self.Einbauwände_Lvl_1 = ""
        self.Einbauwände_Lvl_2 = ""

        self.Türzargen_Lvl_1 = ""
        self.Türzargen_Lvl_2 = ""

        self.Türen_Lvl_1 = ""
        self.Türen_Lvl_2 = ""

        self.Malerarbeiten_Anstriche_Lvl_1 = ""
        self.Malerarbeiten_Anstriche_Lvl_2 = ""

        self.Malerarbeiten_Spachtelarbeiten_Lvl_1 = ""
        self.Malerarbeiten_Spachtelarbeiten_Lvl_2 = ""

        self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = ""
        self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = ""

        self.Mengen_Oberflächen_Unterschicht_Lvl_1 = ""
        self.Mengen_Oberflächen_Unterschicht_Lvl_2 = ""

        self.Mengen_Oberflächen_Deckschicht_Lvl_1 = ""
        self.Mengen_Oberflächen_Deckschicht_Lvl_2 = ""

    # Walls - fill the Strabag parameters
        # Compare category IDs and get cpiFitMatchKey parameter
        if revit_element.Category.Id == DB.Category.GetCategory(
            self.doc,
                DB.BuiltInCategory.OST_Walls).Id:

            self.cpiFitMatchKey = type_parameter_value_for_instance(
                self.doc,
                revit_element.Id,
                "cpiFitMatchKey")
            if self.cpiFitMatchKey is not None:

                # Analyse parameters of Revit element, to assign Strabag parameters
            # Wandbeläge
                if self.cpiFitMatchKey == "5D-AB-TF-IW-FL":
                    self.Wandbeläge_Lvl_1 = "Wandfliesen"
                    self.Wandbeläge_Lvl_2 = "Wandfliesen"
                elif self.cpiFitMatchKey == "5D-AB-TF-IW-NAS":
                    self.Wandbeläge_Lvl_1 = "Natursteinbeläge Wand"
                    self.Wandbeläge_Lvl_2 = "Natursteinbeläge Wand"
            # Innenputz
                # work only with elements with cpiFitMatchKey "5D-AB-TF-IW"
                if self.cpiFitMatchKey == "5D-AB-TF-IW":
                    # get special inastance parameters of Revit element
                    Wand_Unterschicht_auf_Ortbeton = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Ortbeton').\
                        AsString()
                    Wand_Unterschicht_auf_Fertigteile = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Fertigteile').\
                        AsString()
                    Wand_Unterschicht_auf_Mauerwerk = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Mauerwerk').\
                        AsString()
                    genElementMatchKey = self.revit_element.\
                        LookupParameter('5dgen_genElementMatchKey').\
                        AsString()

                    # three conditions for the same Strabag parameters but with
                    # different revit parameters:
                    if genElementMatchKey is not None:  # some w have no this par.
                        if Wand_Unterschicht_auf_Ortbeton is not None:  # some w have no this par.
                            if Wand_Unterschicht_auf_Ortbeton.find("putz") != -1\
                                and genElementMatchKey.find("5D-KI-STB-OB-") == 0\
                                    and not genElementMatchKey.find(
                                        "5D-KI-STB-OB-TR") == 0:
                                self.Innenputz_Lvl_1 = "Innenputz"
                                self.Innenputz_Lvl_2 = "Innenputzarbeiten Wand"

                        if Wand_Unterschicht_auf_Fertigteile is not None:  # some w have no this par.
                            if Wand_Unterschicht_auf_Fertigteile.find("putz") != -1\
                                and genElementMatchKey.find("5D-KI-STB-FT-") == 0\
                                    and not genElementMatchKey.find(
                                        "5D-KI-STB-OB-TR") == 0:
                                self.Innenputz_Lvl_1 = "Innenputz"
                                self.Innenputz_Lvl_2 = "Innenputzarbeiten Wand"

                        if Wand_Unterschicht_auf_Mauerwerk is not None:  # some w have no this par.
                            if Wand_Unterschicht_auf_Mauerwerk.find("putz") != -1\
                                    and genElementMatchKey.find("5D-KI-MW-") == 0:
                                self.Innenputz_Lvl_1 = "Innenputz"
                                self.Innenputz_Lvl_2 = "Innenputzarbeiten Wand"
            # Trockenbau / Verkleidungen Wand
                if self.cpiFitMatchKey.find("5D-AB-IW-GK") == 0:
                    self.Trockenbau_Verkleidungen_Wand_Lvl_1 = "Trockenbau / Verkleidungen Wand"
                    self.Trockenbau_Verkleidungen_Wand_Lvl_2 = "Gipskartonständerwand"
                elif self.cpiFitMatchKey.find("5D-AB-IW-GWBP-") == 0:
                    self.Trockenbau_Verkleidungen_Wand_Lvl_1 = "Trockenbau / Verkleidungen Wand"
                    self.Trockenbau_Verkleidungen_Wand_Lvl_2 = "Gipsplattenwände"
                elif self.cpiFitMatchKey.find("5D-AB-IW-WV") == 0:
                    self.Trockenbau_Verkleidungen_Wand_Lvl_1 = "Trockenbau / Verkleidungen Wand"
                    self.Trockenbau_Verkleidungen_Wand_Lvl_2 = "Wandverkleidungen"
            # Einbauwände
                if self.cpiFitMatchKey.find("5D-AB-IW-EB-G") == 0:
                    self.Einbauwände_Lvl_1 = "Einbauwände"
                    self.Einbauwände_Lvl_2 = "Systemtrennwände Glas"
                elif self.cpiFitMatchKey.find("5D-AB-IW-EB-VE") == 0:
                    self.Einbauwände_Lvl_1 = "Einbauwände"
                    self.Einbauwände_Lvl_2 = "Systemtrennwände Vollwandelement"
                elif self.cpiFitMatchKey == "5D-AB-IW-EB-WC":
                    self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                    self.Einbauwände_Lvl_2 = "WC-Trennwand"
                elif self.cpiFitMatchKey == "5D-AB-IW-EB-MTN-H":
                    self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                    self.Einbauwände_Lvl_2 = "Mobile Schiebewand Holz"
                elif self.cpiFitMatchKey == "5D-AB-IW-EB-MTN-G":
                    self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                    self.Einbauwände_Lvl_2 = "Glas-Schiebewand"
                elif self.cpiFitMatchKey.find("5D-AB-IW-EB-KT") == 0:
                    self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                    self.Einbauwände_Lvl_2 = "Kellertrennwand"
            # Allgemeiner Ausbau and Nicht zugeordnete Mengen Oberflächen
                if self.cpiFitMatchKey == "5D-AB-TF-IW":
                # get special inastance parameters of Revit element
                    genElementMatchKey = self.revit_element.\
                        LookupParameter('5dgen_genElementMatchKey').\
                        AsString()

                    # Anstriche Wand parameters
                    Wand_Deckschicht = self.revit_element.\
                        LookupParameter('5d Wand Deckschicht').\
                        AsString()
                    Wand_Zwischenschicht = self.revit_element.\
                        LookupParameter('5d Wand Zwischenschicht').\
                        AsString()

                    # Spachtelarbeiten Wand parameters
                    Wand_Unterschicht_auf_Fertigteile = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Fertigteile').\
                        AsString()
                    Wand_Unterschicht_auf_Ortbeton = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Ortbeton').\
                        AsString()
                    Wand_Unterschicht_auf_Mauerwerk = self.revit_element.\
                        LookupParameter('5d Wand Unterschicht auf Mauerwerk').\
                        AsString()

                # Anstriche Wand
                    if Wand_Deckschicht is not None\
                            and Wand_Zwischenschicht is not None\
                            and genElementMatchKey is not None:  # some w have no this par.

                        if genElementMatchKey.find("5D-KI-STB-OB-") == 0:
                            if (Wand_Deckschicht.find("Dispersion") == 0
                                    and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1)\
                                    or (Wand_Zwischenschicht == "Kunstharzputz"
                                        and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1):
                                self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                        elif genElementMatchKey.find("5D-KI-STB-FT-") == 0:
                            if (Wand_Deckschicht.find("Dispersion") == 0
                                    and genElementMatchKey.find("5D-KI-STB-FT-TR") == -1)\
                                    or (Wand_Zwischenschicht == "Kunstharzputz"
                                        and genElementMatchKey.find("5D-KI-STB-FT-TR") == -1):
                                self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                        elif genElementMatchKey.find("5D-KI-MW-") == 0:
                            if Wand_Deckschicht.find("Dispersion") == 0\
                                    or Wand_Zwischenschicht == "Kunstharzputz":
                                self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                        elif genElementMatchKey.find("5D-AB-IW-GK") == 0:
                            if Wand_Deckschicht.find("Dispersion") == 0\
                                    or Wand_Zwischenschicht == "Kunstharzputz":
                                self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                        elif genElementMatchKey.find("5D-AB-IW-GWBP") == 0:
                            if Wand_Deckschicht.find("Dispersion") == 0\
                                    or Wand_Zwischenschicht == "Kunstharzputz":
                                self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                        elif genElementMatchKey.find("5D-KI-WD") == 0\
                                and Wand_Deckschicht.find("Dispersion") == 0:
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Wand"

                # Spachtelarbeiten Wand (Spachtelung)
                    if Wand_Unterschicht_auf_Fertigteile is not None\
                            and Wand_Unterschicht_auf_Ortbeton is not None\
                            and genElementMatchKey is not None:  # some w have no this par.

                        if genElementMatchKey.find("5D-KI-STB-OB-") == 0:
                            if Wand_Unterschicht_auf_Ortbeton.find("spachtelung") != -1\
                                    and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1:
                                self.Malerarbeiten_Spachtelarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Spachtelarbeiten_Lvl_2 = "Spachtelarbeiten Wand"

                        elif genElementMatchKey.find("5D-KI-STB-FT-") == 0:
                            if Wand_Unterschicht_auf_Fertigteile.find("spachtelung") != -1\
                                    and genElementMatchKey.find("5D-KI-STB-FT-TR") == -1:
                                self.Malerarbeiten_Spachtelarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Spachtelarbeiten_Lvl_2 = "Spachtelarbeiten Wand"

                # Tapezierarbeiten Wand
                    if Wand_Zwischenschicht is not None\
                            and genElementMatchKey is not None:  # some w have no this par.

                        if genElementMatchKey.find("5D-KI-STB-OB-") == 0:
                            if (Wand_Zwischenschicht.find("vlies") != -1
                                    and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1)\
                                    or (Wand_Zwischenschicht.find("tapete") != -1
                                        and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1):
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = "Tapezierarbeiten Wand"

                        elif genElementMatchKey.find("5D-KI-STB-FT-") == 0:
                            if (Wand_Zwischenschicht.find("vlies") != -1
                                    and genElementMatchKey.find("5D-KI-STB-FT-TR") == -1)\
                                    or (Wand_Zwischenschicht.find("tapete") != -1
                                        and genElementMatchKey.find("5D-KI-STB-OB-TR") == -1):
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = "Tapezierarbeiten Wand"

                        elif genElementMatchKey.find("5D-KI-MW-") == 0:
                            if Wand_Zwischenschicht.find("vlies") != -1\
                                    or Wand_Zwischenschicht.find("tapete") != -1:
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = "Tapezierarbeiten Wand"

                        elif genElementMatchKey.find("5D-AB-IW-GK") == 0:
                            if Wand_Zwischenschicht.find("vlies") != -1\
                                    or Wand_Zwischenschicht.find("tapete") != -1:
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = "Tapezierarbeiten Wand"

                        elif genElementMatchKey.find("5D-AB-IW-GWBP") == 0:
                            if Wand_Zwischenschicht.find("vlies") != -1\
                                    or Wand_Zwischenschicht.find("tapete") != -1:
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_1 = "Malerarbeiten"
                                self.Malerarbeiten_Tapezierarbeiten_Lvl_2 = "Tapezierarbeiten Wand"

                # Mengen Wand Unterschicht
                    if self.Innenputz_Lvl_2 == "" and self.Malerarbeiten_Spachtelarbeiten_Lvl_2 == "":  # delete G40 and G91 positions from summary schedule  (excel) 

                        # if Wand_Unterschicht_auf_Fertigteile is not None\
                        #         and Wand_Unterschicht_auf_Ortbeton is not None\
                        #         and Wand_Unterschicht_auf_Mauerwerk is not None\
                        #         and genElementMatchKey is not None:  # some w have no this par.

                        if genElementMatchKey is not None:  # some w have no this par.

                            # first attempt
                                # if genElementMatchKey.find("5D-KI-STB-OB-") == 0:
                                #     if len(Wand_Unterschicht_auf_Ortbeton) > 0:
                                #         if Wand_Unterschicht_auf_Ortbeton != "-"\
                                #                 or genElementMatchKey.find("5D-KI-STB-OB-TR") != 0:
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"
                                #     elif Wand_Unterschicht_auf_Ortbeton != "-"\
                                #             and genElementMatchKey.find("5D-KI-STB-OB-TR") == 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                # elif genElementMatchKey.find("5D-KI-STB-FT-") == 0:
                                #     if len(Wand_Unterschicht_auf_Fertigteile) > 0:
                                #         if Wand_Unterschicht_auf_Fertigteile != "-"\
                                #                 or genElementMatchKey.find("5D-KI-STB-FT-TR") != 0:
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"
                                #     elif Wand_Unterschicht_auf_Fertigteile != "-"\
                                #             and genElementMatchKey.find("5D-KI-STB-FT-TR") == 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                # elif genElementMatchKey.find("5D-KI-MW-") == 0:
                                #     if len(Wand_Unterschicht_auf_Mauerwerk) > 0\
                                #             and Wand_Unterschicht_auf_Mauerwerk != "-":
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                # if genElementMatchKey.find("5D-KI-STB-OB-") == 0:
                                #     if len(Wand_Unterschicht_auf_Ortbeton) > 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                #     if Wand_Unterschicht_auf_Ortbeton == "-"\
                                #             or genElementMatchKey.find("5D-KI-STB-OB-TR") == 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = ""
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = ""

                                #     if genElementMatchKey.find("5D-KI-STB-OB-TR") == 0:
                                #         if Wand_Unterschicht_auf_Ortbeton == "-":
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                # if genElementMatchKey.find("5D-KI-STB-FT-") == 0:
                                #     if len(Wand_Unterschicht_auf_Fertigteile) > 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                #     if Wand_Unterschicht_auf_Fertigteile == "-"\
                                #             or genElementMatchKey.find("5D-KI-STB-FT-TR") == 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = ""
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = ""

                                #     if genElementMatchKey.find("5D-KI-STB-FT-TR") == 0:
                                #         if Wand_Unterschicht_auf_Fertigteile == "-":
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #             self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                # if genElementMatchKey.find("5D-KI-MW-") == 0:
                                #     if len(Wand_Unterschicht_auf_Mauerwerk) > 0:
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                #     if Wand_Unterschicht_auf_Mauerwerk == "-":
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_1 = ""
                                #         self.Mengen_Oberflächen_Unterschicht_Lvl_2 = ""

                            if (genElementMatchKey.find("5D-KI-STB-OB-") == 0 and len(Wand_Unterschicht_auf_Ortbeton) > 0)\
                                    or (genElementMatchKey.find("5D-KI-STB-FT-") == 0 and len(Wand_Unterschicht_auf_Fertigteile) > 0)\
                                    or (genElementMatchKey.find("5D-KI-MW-") == 0 and len(Wand_Unterschicht_auf_Mauerwerk) > 0):
                                self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                                if (genElementMatchKey.find("5D-KI-STB-OB-TR") == 0 and len(Wand_Unterschicht_auf_Ortbeton) > 0)\
                                        or (genElementMatchKey.find("5D-KI-STB-FT-TR") == 0 and len(Wand_Unterschicht_auf_Fertigteile) > 0)\
                                        or (genElementMatchKey.find("5D-KI-STB-OB-") == 0 and Wand_Unterschicht_auf_Ortbeton == "-")\
                                        or (genElementMatchKey.find("5D-KI-STB-FT-") == 0 and Wand_Unterschicht_auf_Ortbeton == "-")\
                                        or (genElementMatchKey.find("5D-KI-MW-") == 0 and Wand_Unterschicht_auf_Ortbeton == "-"):
                                    self.Mengen_Oberflächen_Unterschicht_Lvl_1 = ""
                                    self.Mengen_Oberflächen_Unterschicht_Lvl_2 = ""

                                    if (genElementMatchKey.find("5D-KI-STB-OB-TR") == 0 and Wand_Unterschicht_auf_Ortbeton == "-")\
                                            or (genElementMatchKey.find("5D-KI-STB-FT-TR") == 0 and Wand_Unterschicht_auf_Fertigteile == "-"):
                                        self.Mengen_Oberflächen_Unterschicht_Lvl_1 = "Mengen_Oberflächen"
                                        self.Mengen_Oberflächen_Unterschicht_Lvl_2 = "Mengen Wand Unterschicht"

                # Mengen Wand Deckschicht
                    if self.Malerarbeiten_Anstriche_Lvl_1 == "":  # delete G89 position from summary schedule  (excel) 
                        if genElementMatchKey is not None:  # some w have no this par.
                            if (genElementMatchKey.find("5D-KI-STB-OB-") == 0 and len(Wand_Deckschicht) > 0)\
                                    or (genElementMatchKey.find("5D-KI-STB-FT-") == 0 and len(Wand_Zwischenschicht) > 0)\
                                    or (genElementMatchKey.find("5D-KI-MW-") == 0 and len(Wand_Deckschicht) > 0)\
                                    or (genElementMatchKey.find("5D-AB-IW-GK") == 0 and len(Wand_Zwischenschicht) > 0)\
                                    or (genElementMatchKey.find("5D-AB-IW-GWBP") == 0 and len(Wand_Deckschicht) > 0):
                                self.Mengen_Oberflächen_Deckschicht_Lvl_1 = "Mengen_Oberflächen"
                                self.Mengen_Oberflächen_Deckschicht_Lvl_2 = "Mengen Wand Deckschicht"

                                if (genElementMatchKey.find("5D-KI-STB-OB-TR") == 0 and len(Wand_Deckschicht) > 0)\
                                        or (genElementMatchKey.find("5D-KI-STB-FT-TR") == 0 and len(Wand_Zwischenschicht) > 0)\
                                        or self.Malerarbeiten_Anstriche_Lvl_1 != ""\
                                        or (genElementMatchKey.find("5D-KI-STB-OB-") == 0 and Wand_Deckschicht == "-")\
                                        or (genElementMatchKey.find("5D-KI-STB-FT-") == 0 and Wand_Deckschicht == "-")\
                                        or (genElementMatchKey.find("5D-KI-MW-") == 0 and Wand_Deckschicht == "-")\
                                        or (genElementMatchKey.find("5D-AB-IW-GK") == 0 and Wand_Deckschicht == "-")\
                                        or (genElementMatchKey.find("5D-AB-IW-GWBP") == 0 and Wand_Deckschicht == "-"):
                                    self.Mengen_Oberflächen_Deckschicht_Lvl_1 = ""
                                    self.Mengen_Oberflächen_Deckschicht_Lvl_2 = ""

                                    if (genElementMatchKey.find("5D-KI-STB-OB-TR") == 0 and Wand_Deckschicht == "-")\
                                            or (genElementMatchKey.find("5D-KI-STB-FT-TR") == 0 and Wand_Deckschicht == "-")\
                                            or (genElementMatchKey.find("5D-KI-STB-OB-") == 0 and Wand_Zwischenschicht == "Kunstharzputz")\
                                            or (genElementMatchKey.find("5D-KI-STB-FT-") == 0 and Wand_Zwischenschicht == "Kunstharzputz")\
                                            or (genElementMatchKey.find("5D-KI-MW-") == 0 and Wand_Zwischenschicht == "Kunstharzputz")\
                                            or (genElementMatchKey.find("5D-AB-IW-GK") == 0 and Wand_Zwischenschicht == "Kunstharzputz")\
                                            or (genElementMatchKey.find("5D-AB-IW-GWBP") == 0 and Wand_Zwischenschicht == "Kunstharzputz"):
                                        self.Mengen_Oberflächen_Deckschicht_Lvl_1 = "Mengen_Oberflächen"
                                        self.Mengen_Oberflächen_Deckschicht_Lvl_2 = "Mengen Wand Deckschicht"

                                        if (genElementMatchKey.find("5D-KI-STB-OB-TR") == 0 and Wand_Zwischenschicht == "Kunstharzputz")\
                                                or (genElementMatchKey.find("5D-KI-STB-FT-TR") == 0 and Wand_Zwischenschicht == "Kunstharzputz"):
                                            self.Mengen_Oberflächen_Deckschicht_Lvl_1 = ""
                                            self.Mengen_Oberflächen_Deckschicht_Lvl_2 = ""

                                            if (genElementMatchKey.find("5D-KI-WD") == 0 and len(Wand_Deckschicht) > 0):
                                                self.Mengen_Oberflächen_Deckschicht_Lvl_1 = "Mengen_Oberflächen"
                                                self.Mengen_Oberflächen_Deckschicht_Lvl_2 = "Mengen Wand Deckschicht"

    # Floors - fill the Strabag parameters
        if revit_element.Category.Id == DB.Category.GetCategory(
            self.doc,
                DB.BuiltInCategory.OST_Floors).Id:

            # get Revit element parameters
            self.cpiFitMatchKey = type_parameter_value_for_instance(
                self.doc,
                revit_element.Id,
                "cpiFitMatchKey")

            Fußbodenbelag = type_parameter_value_for_instance(
                self.doc,
                revit_element.Id,
                "5d Fußbodenbelag")

            # Analyse parameters of Revit element, to assign Strabag parameters
        # Estriche_Bodensysteme
            if self.cpiFitMatchKey.find("5D-AB-BK-CT") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Estriche/Bodensysteme_Lvl_3_Estriche_Lvl_4_Zementestrich"
                self.Estriche_Bodensysteme_Lvl_1 = "Estriche"
                self.Estriche_Bodensysteme_Lvl_2 = "Zementestrich"
            elif self.cpiFitMatchKey.find("5D-AB-BK-CA") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Estriche/Bodensysteme_Lvl_3_Estriche_Lvl_4_Calciumsulfat_Estrich Wand"
                self.Estriche_Bodensysteme_Lvl_1 = "Estriche"
                self.Estriche_Bodensysteme_Lvl_2 = "Calciumsulfat_Estrich Wand"
            elif self.cpiFitMatchKey.find("5D-AB-BK-GA") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Estriche/Bodensysteme_Lvl_3_Estriche_Lvl_4_Gussasphaltestrich Wand"
                self.Estriche_Bodensysteme_Lvl_1 = "Estriche"
                self.Estriche_Bodensysteme_Lvl_2 = "Gussasphaltestrich Wand"
        # Hohlraumböden and Doppelböden
            if self.cpiFitMatchKey.find("5D-AB-BK-HoBo") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Estriche/Bodensysteme_Lvl_3_Hohlraumböden_Lvl_4_Hohlraumböden"
                self.Estriche_Bodensysteme_Lvl_1 = "Hohlraumböden"
                self.Estriche_Bodensysteme_Lvl_2 = "Hohlraumböden"
            elif self.cpiFitMatchKey.find("5D-AB-BK-DoBo") == 0:
                self.summary += "_Lvl_1_Innenausbau_Lvl_2_Estriche/Bodensysteme_Lvl_3_Doppelböden_Lvl_4_Doppelböden"
                self.Estriche_Bodensysteme_Lvl_1 = "Doppelböden"
                self.Estriche_Bodensysteme_Lvl_2 = "Doppelböden"

            if Fußbodenbelag is not None:  # some floors have no this parameter

        # Bodenbeläge
                if self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.10") == 0:
                    self.Bodenbeläge_Lvl_1 = "Natursteinbeläge Boden"
                    self.Bodenbeläge_Lvl_2 = "Natursteinbeläge Boden"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.20") == 0:
                    self.Bodenbeläge_Lvl_1 = "Betonwerksteinbeläge"
                    self.Bodenbeläge_Lvl_2 = "Betonwerksteinbeläge"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.31") == 0:
                    self.Bodenbeläge_Lvl_1 = "Fliesenbeläge"
                    self.Bodenbeläge_Lvl_2 = "Bodenfliesen"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.32") == 0:
                    self.Bodenbeläge_Lvl_1 = "Fliesenbeläge"
                    self.Bodenbeläge_Lvl_2 = "Rüttelboden"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.41") == 0:
                    self.Bodenbeläge_Lvl_1 = "Textile / elastische_Beläge"
                    self.Bodenbeläge_Lvl_2 = "Textile Beläge"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.42") == 0:
                    self.Bodenbeläge_Lvl_1 = "Textile / elastische_Beläge"
                    self.Bodenbeläge_Lvl_2 = "Kunststoffbeläge"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.43") == 0:
                    self.Bodenbeläge_Lvl_1 = "Textile / elastische_Beläge"
                    self.Bodenbeläge_Lvl_2 = "Linoleumbeläge"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.44") == 0:
                    self.Bodenbeläge_Lvl_1 = "Textile / elastische_Beläge"
                    self.Bodenbeläge_Lvl_2 = "Gummibeläge"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.51") == 0:
                    self.Bodenbeläge_Lvl_1 = "Holzböden"
                    self.Bodenbeläge_Lvl_2 = "Parkettböden"
                elif self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.33.52") == 0:
                    self.Bodenbeläge_Lvl_1 = "Holzböden"
                    self.Bodenbeläge_Lvl_2 = "Parkettböden"

        # Oberflächenschutz Estrich / Beton
                if self.cpiFitMatchKey.find("5D-AB-BK-") == 0\
                        and Fußbodenbelag.find("6.32.10.") == 0:
                    if Fußbodenbelag.find("6.32.10.5") == 0\
                            or Fußbodenbelag.find("6.32.10.6")== 0:
                        self.Oberflächenschutz_Estrich_Beton_Lvl_1 = "Oberflächenschutz Estrich / Beton"
                        self.Oberflächenschutz_Estrich_Beton_Lvl_2 = "Boden-Beschichtungen"
                    elif Fußbodenbelag.find("6.32.10.4") == 0\
                            or Fußbodenbelag.find("6.32.10.3")== 0:
                        self.Oberflächenschutz_Estrich_Beton_Lvl_1 = "Oberflächenschutz Estrich / Beton"
                        self.Oberflächenschutz_Estrich_Beton_Lvl_2 = "Boden-Versiegelungen"
                    elif Fußbodenbelag.find("6.32.10.2") == 0:
                        self.Oberflächenschutz_Estrich_Beton_Lvl_1 = "Oberflächenschutz Estrich / Beton"
                        self.Oberflächenschutz_Estrich_Beton_Lvl_2 = "Boden-Anstriche "

        # Innenputz
            # work only with elements with cpiFitMatchKey "5D-AB-TF-DE"
            if self.cpiFitMatchKey == "5D-AB-TF-DE":
                # get special inastance parameters of Revit element
                Decken_Unterschicht = self.revit_element.\
                    LookupParameter('5d Decken Unterschicht').\
                    AsString()
                genElementMatchKey = self.revit_element.\
                    LookupParameter('5dgen_genElementMatchKey').\
                    AsString()

                if genElementMatchKey is not None and Decken_Unterschicht is not None:  # some w have no this par.
                    if Decken_Unterschicht.find("putz") != -1:
                        if genElementMatchKey.find("5D-KI-STB-OB-DE") == 0\
                                or genElementMatchKey.find("5D-KI-STB-OB-TR-") == 0\
                                or genElementMatchKey.find("5D-KI-STB-FT-DE") == 0\
                                or genElementMatchKey.find("5D-KI-STB-FT-TR-") == 0:
                            self.Innenputz_Lvl_1 = "Innenputz"
                            self.Innenputz_Lvl_2 = "Innenputzarbeiten Decke"

        # Allgemeiner Ausbau and Nicht zugeordnete Mengen Oberflächen:  
            if self.cpiFitMatchKey == "5D-AB-TF-DE":
                # get special inastance parameters of Revit element
                genElementMatchKey = self.revit_element.\
                    LookupParameter('5dgen_genElementMatchKey').\
                    AsString()
                # Anstriche Decke parameters
                Decken_Deckschicht = self.revit_element.\
                    LookupParameter('5d Decken Deckschicht').\
                    AsString()
                Decken_Zwischenschicht = self.revit_element.\
                    LookupParameter('5d Decken Zwischenschicht').\
                    AsString()
                # Spachtelarbeiten Decke parameters
                Decken_Unterschicht = self.revit_element.\
                    LookupParameter('5d Decken Unterschicht').\
                    AsString()

            # Anstriche Decke
                if Decken_Deckschicht is not None\
                        and Decken_Zwischenschicht is not None\
                        and genElementMatchKey is not None:  # some f have no this par.

                    if genElementMatchKey.find("5D-KI-STB-OB-DE") == 0:
                        if Decken_Deckschicht.find("Dispersion") == 0\
                                or Decken_Zwischenschicht == "Kunstharzputz":
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

                    elif genElementMatchKey.find("5D-KI-STB-OB-TR-") == 0:
                        if Decken_Deckschicht.find("Dispersion") == 0\
                                or Decken_Zwischenschicht == "Kunstharzputz":
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

                    elif genElementMatchKey.find("5D-KI-STB-FT-DE") == 0:
                        if Decken_Deckschicht.find("Dispersion") == 0\
                                or Decken_Zwischenschicht == "Kunstharzputz":
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

                    elif genElementMatchKey.find("5D-KI-STB-FT-TR-") == 0:
                        if Decken_Deckschicht.find("Dispersion") == 0\
                                or Decken_Zwischenschicht == "Kunstharzputz":
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

                    elif genElementMatchKey.find("5D-AB-AHD-GK") == 0:
                        if Decken_Deckschicht.find("Dispersion") == 0\
                                or Decken_Zwischenschicht == "Kunstharzputz":
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

                    elif genElementMatchKey.find("5D-KI-WD") == 0\
                            and Decken_Deckschicht.find("Dispersion") == 0:
                        self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                        self.Malerarbeiten_Anstriche_Lvl_2 = "Anstriche Decke"

            # Spachtelarbeiten Decke
                if Decken_Unterschicht is not None\
                        and genElementMatchKey is not None:  # some f have no this par.

                    if genElementMatchKey.find("5D-KI-STB-OB-DE") == 0\
                            or genElementMatchKey.find("5D-KI-STB-OB-TR-") == 0\
                            or genElementMatchKey.find("5D-KI-STB-FT-DE") == 0\
                            or genElementMatchKey.find("5D-KI-STB-FT-TR-") == 0:
                        if Decken_Unterschicht.find("spachtelung") != -1:
                            self.Malerarbeiten_Spachtelarbeiten_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Spachtelarbeiten_Lvl_2 = "Spachtelarbeiten Decke"

            # Tapezierarbeiten Decke
                if Decken_Zwischenschicht is not None\
                        and genElementMatchKey is not None:  # some f have no this par.

                    if genElementMatchKey.find("5D-KI-STB-OB-DE") == 0\
                            or genElementMatchKey.find("5D-KI-STB-OB-TR") == 0\
                            or genElementMatchKey.find("5D-KI-STB-FT-DE") == 0\
                            or genElementMatchKey.find("5D-KI-STB-FT-TR") == 0\
                            or genElementMatchKey.find("5D-AB-AHD-GK") == 0:
                        if Decken_Zwischenschicht.find("tapete") != -1\
                                or Decken_Zwischenschicht.find("vlies") != -1:
                            self.Malerarbeiten_Anstriche_Lvl_1 = "Malerarbeiten"
                            self.Malerarbeiten_Anstriche_Lvl_2 = "Tapezierarbeiten Decke"

            # Mengen Decke Deckschicht
                genElementMatchKey_list = [
                    "5D-KI-STB-OB-DE",
                    "5D-KI-STB-OB-TR-",
                    "5D-KI-STB-FT-DE",
                    "5D-KI-STB-FT-TR-",
                    "5D-AB-AHD-GK"]

                if self.Malerarbeiten_Anstriche_Lvl_1 == "":  # delete G90 position from summary schedule  (excel) 
                    if genElementMatchKey is not None:  # some floors have no this par.
                        for MK in genElementMatchKey_list:
                            if genElementMatchKey.find(MK) == 0 and len(Decken_Deckschicht) > 0:
                                self.Mengen_Oberflächen_Deckschicht_Lvl_1 = "Mengen_Oberflächen"
                                self.Mengen_Oberflächen_Deckschicht_Lvl_2 = "Mengen Decke Deckschicht"

                                if Decken_Deckschicht == "Textiltapete":
                                    self.Mengen_Oberflächen_Deckschicht_Lvl_1 = ""
                                    self.Mengen_Oberflächen_Deckschicht_Lvl_2 = ""

                                    if genElementMatchKey.find("5D-KI-WD") == 0 and len(Decken_Deckschicht) > 0:
                                        self.Mengen_Oberflächen_Deckschicht_Lvl_1 = "Mengen_Oberflächen"
                                        self.Mengen_Oberflächen_Deckschicht_Lvl_2 = "Mengen Decke Deckschicht"

    # Ceilings - fill the Strabag parameters
        if revit_element.Category.Id == DB.Category.GetCategory(
            self.doc,
                DB.BuiltInCategory.OST_Ceilings).Id:

            self.cpiFitMatchKey = type_parameter_value_for_instance(
                self.doc,
                revit_element.Id,
                "cpiFitMatchKey")

            # Trockenbau_Verkleidungen_Decke
            if self.cpiFitMatchKey.find("5D-AB-AHD-GK") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Trockenbau/Verkleidungen Decke_Lvl_3_Unterdecken_Lvl_4_Gipskartondecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_1 = "Unterdecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_2 = "Gipskartondecken"
            elif self.cpiFitMatchKey.find("5D-AB-AHD-MF") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Trockenbau/Verkleidungen Decke_Lvl_3_Unterdecken_Lvl_4_Mineralwolldecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_1 = "Unterdecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_2 = "Mineralwolldecken"
            elif self.cpiFitMatchKey.find("5D-AB-AHD-ME") == 0:
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Trockenbau/Verkleidungen Decke_Lvl_3_Unterdecken_Lvl_4_Metalldecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_1 = "Unterdecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_2 = "Metalldecken"
            elif self.cpiFitMatchKey == "5D-AB-AHD-DSL":
                # self.summary += "_Lvl_1_Innenausbau_Lvl_2_Trockenbau/Verkleidungen Decke_Lvl_3_Unterdecken_Lvl_4_Deckensegel"
                self.Trockenbau_Verkleidungen_Decke_Lvl_1 = "Unterdecken"
                self.Trockenbau_Verkleidungen_Decke_Lvl_2 = "Deckensegel"

    # Doors - fill the Strabag parameters
        # Compare category IDs and get cpiFitMatchKey parameter
        if revit_element.Category.Id == DB.Category.GetCategory(
            self.doc,
                DB.BuiltInCategory.OST_Doors).Id:

            self.cpiFitMatchKey = type_parameter_value_for_instance(
                self.doc,
                revit_element.Id,
                "cpiFitMatchKey")
        # Analyse parameters of Revit element, to assign Strabag parameters
        if self.cpiFitMatchKey is not None:  # some doors have no this par.
            if self.cpiFitMatchKey == "5D-AB-TU-XX":
                self.Türen_Lvl_1 = "Sondertüren"
                self.Türen_Lvl_2 = "Sondertüren"
            elif self.cpiFitMatchKey.find("5D-AB-TU-SHT") == 0:
                self.Türen_Lvl_1 = "Schiebetüren"
                self.Türen_Lvl_2 = "Schiebetüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-G":
                self.Türen_Lvl_1 = "Ganzglastüren"
                self.Türen_Lvl_2 = "Ganzglastüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-HR":
                self.Türen_Lvl_1 = "Rahmentüren"
                self.Türen_Lvl_2 = "Holzrahmentüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-ALU":
                self.Türen_Lvl_1 = "Rahmentüren"
                self.Türen_Lvl_2 = "Alurahmentüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-STR":
                self.Türen_Lvl_1 = "Rahmentüren"
                self.Türen_Lvl_2 = "Stahlrahmentüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-ALUB":
                self.Türen_Lvl_1 = "Blechtüren"
                self.Türen_Lvl_2 = "Aluminiumstahlblechtüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-EDB":
                self.Türen_Lvl_1 = "Blechtüren"
                self.Türen_Lvl_2 = "Edelstahlblechtüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-SB":
                self.Türen_Lvl_1 = "Blechtüren"
                self.Türen_Lvl_2 = "Stahlblechtüren"
            elif self.cpiFitMatchKey == "5D-AB-TU-H":
                self.Türen_Lvl_1 = "Holztüren"
                self.Türen_Lvl_2 = "Holztüren"
        # Einbauwände
            if self.cpiFitMatchKey == "5D-AB-TU-EB-ALU":
                self.Einbauwände_Lvl_1 = "Einbauwände"
                self.Einbauwände_Lvl_2 = "Systemtrennwandtüren Alurahmen"
            elif self.cpiFitMatchKey == "5D-AB-TU-EB-G":
                self.Einbauwände_Lvl_1 = "Einbauwände"
                self.Einbauwände_Lvl_2 = "Systemtrennwandtüren Glas"
            elif self.cpiFitMatchKey == "5D-AB-TU-EB-H":
                self.Einbauwände_Lvl_1 = "Einbauwände"
                self.Einbauwände_Lvl_2 = "Systemtrennwandtüren Holz"
            elif self.cpiFitMatchKey == "5D-AB-TU-EB-STR":
                self.Einbauwände_Lvl_1 = "Einbauwände"
                self.Einbauwände_Lvl_2 = "Systemtrennwandtüren Stahlrahmen"
            elif self.cpiFitMatchKey == "5D-AB-TU-WC":
                self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                self.Einbauwände_Lvl_2 = "WC-Türen"
            elif self.cpiFitMatchKey == "5D-AB-TU-MTN-H":
                self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                self.Einbauwände_Lvl_2 = "mobile Trennwandtüren Holz"
            elif self.cpiFitMatchKey == "5D-AB-TU-MTN-G":
                self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                self.Einbauwände_Lvl_2 = "mobile Trennwandtüren Glas"
            elif self.cpiFitMatchKey.find("5D-AB-TU-KT") == 0:
                self.Einbauwände_Lvl_1 = "Trennwandanlagen"
                self.Einbauwände_Lvl_2 = "Kelltertüren"


class PBI_collector():
    """ Selection and working with PBI_elements """

    def __init__(self, doc):
        self.doc = doc

    def _select_PBI_elements(self):
        """ Create PBI_elements for the Revit walls and floors. For testing """
        walls = list(
            DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType())

        floors = list(
            DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType())

        ceilings = list(
            DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType())

        doors = list(
            DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType())

        PBI_walls = [PBI_element(doc, wall) for wall in walls]
        PBI_floors = [PBI_element(doc, floor) for floor in floors]
        PBI_ceilngs = [PBI_element(doc, ceiling) for ceiling in ceilings]
        PBI_doors = [PBI_element(doc, door) for door in doors]

        return flatten([PBI_walls, PBI_floors, PBI_ceilngs, PBI_doors])

    def excel_data(self):
        """ Collect data to excel export """
        PBI_elements = self._select_PBI_elements()

        #  Create names for schedule columns:
        data_list = [[
            'element Id',

            'Estriche_Bodensysteme_Lvl_1',
            'Estriche_Bodensysteme_Lvl_2',

            'Oberflächenschutz_Estrich_Beton_Lvl_1',
            'Oberflächenschutz_Estrich_Beton_Lvl_2',

            'Bodenbeläge_Lvl_1',
            'Bodenbeläge_Lvl_2',

            'Wandbeläge_Lvl_1',
            'Wandbeläge_Lvl_2',

            'Innenputz_Lvl_1',
            'Innenputz_Lvl_2',

            'Trockenbau_Verkleidungen_Wand_Lvl_1',
            'Trockenbau_Verkleidungen_Wand_Lvl_2',

            'Trockenbau_Verkleidungen_Decke_Lvl_1',
            'Trockenbau_Verkleidungen_Decke_Lvl_2',

            'Einbauwände_Lvl_1',
            'Einbauwände_Lvl_2',

            'Türzargen_Lvl_1',
            'Türzargen_Lvl_2',

            'Türen_Lvl_1',
            'Türen_Lvl_2',

            'Malerarbeiten_Anstriche_Lvl_1',
            'Malerarbeiten_Anstriche_Lvl_2',
            'Malerarbeiten_Spachtelarbeiten_Lvl_1',
            'Malerarbeiten_Spachtelarbeiten_Lvl_2',
            'Malerarbeiten_Tapezierarbeiten_Lvl_1',
            'Malerarbeiten_Tapezierarbeiten_Lvl_2',

            'Mengen_Oberflächen_Unterschicht_Lvl_1',
            'Mengen_Oberflächen_Unterschicht_Lvl_2',

            'Mengen_Oberflächen_Deckschicht_Lvl_1',
            'Mengen_Oberflächen_Deckschicht_Lvl_2'
            ]]

        # create data list of every elements with ID and Strabag parameters:
        for el in PBI_elements:
            element_list = []
            element_list.append(el.revit_element.Id)

            element_list.append(el.Estriche_Bodensysteme_Lvl_1)
            element_list.append(el.Estriche_Bodensysteme_Lvl_2)

            element_list.append(el.Oberflächenschutz_Estrich_Beton_Lvl_1)
            element_list.append(el.Oberflächenschutz_Estrich_Beton_Lvl_2)

            element_list.append(el.Bodenbeläge_Lvl_1)
            element_list.append(el.Bodenbeläge_Lvl_2)

            element_list.append(el.Wandbeläge_Lvl_1)
            element_list.append(el.Wandbeläge_Lvl_2)

            element_list.append(el.Innenputz_Lvl_1)
            element_list.append(el.Innenputz_Lvl_2)

            element_list.append(el.Trockenbau_Verkleidungen_Wand_Lvl_1)
            element_list.append(el.Trockenbau_Verkleidungen_Wand_Lvl_2)

            element_list.append(el.Trockenbau_Verkleidungen_Decke_Lvl_1)
            element_list.append(el.Trockenbau_Verkleidungen_Decke_Lvl_2)

            element_list.append(el.Einbauwände_Lvl_1)
            element_list.append(el.Einbauwände_Lvl_2)

            element_list.append(el.Türzargen_Lvl_1)
            element_list.append(el.Türzargen_Lvl_2)

            element_list.append(el.Türen_Lvl_1)
            element_list.append(el.Türen_Lvl_2)

            element_list.append(el.Malerarbeiten_Anstriche_Lvl_1)
            element_list.append(el.Malerarbeiten_Anstriche_Lvl_2)
            element_list.append(el.Malerarbeiten_Spachtelarbeiten_Lvl_1)
            element_list.append(el.Malerarbeiten_Spachtelarbeiten_Lvl_2)
            element_list.append(el.Malerarbeiten_Tapezierarbeiten_Lvl_1)
            element_list.append(el.Malerarbeiten_Tapezierarbeiten_Lvl_2)

            element_list.append(el.Mengen_Oberflächen_Unterschicht_Lvl_1)
            element_list.append(el.Mengen_Oberflächen_Unterschicht_Lvl_2)

            element_list.append(el.Mengen_Oberflächen_Deckschicht_Lvl_1)
            element_list.append(el.Mengen_Oberflächen_Deckschicht_Lvl_2)

            data_list.append(element_list)

        return data_list

    def tests(self):
        doors = list(
            DB.FilteredElementCollector(doc).OfCategory(
                DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType())
        return [PBI_element(doc, door).cpiFitMatchKey for door in doors]


doc = DocumentManager.Instance.CurrentDBDocument
PBI_elements = PBI_collector(doc).excel_data()

OUT = PBI_elements
