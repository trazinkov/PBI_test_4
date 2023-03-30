# -*- coding: utf-8 -*-
# convert Revit id elements into 5 levels of text parameters
# Analyze csv files with ID elements and Strabag parameters
# Important - delete all ',' in titels of csv!!!

'''add lib'''
import sys
sys.path += [
    r"Q:\DESTUALB\15-BIM5D\10-B5D\22-Hochbau\60-ENTWICKLUNG\Dynamo\Templates\Python"]  # noqa

from CSV_table import CSV_table

file = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 11 - Tables records\Excel\Tables_ID_elements - floors.csv"  # noqa
file_walls = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 11 - Tables records\Excel\Tables_ID_elements - walls.csv"  # noqa
file_doors = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 11 - Tables records\Excel\Tables_ID_elements - doors.csv"  # noqa
file_ceilings = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 11 - Tables records\Excel\Tables_ID_elements - ceilings.csv"  # noqa
levels_match = r"C:\Anatoly\20 ZT\22 Power BI  - summary schedule\Development\Phase 11 - Tables records\Excel\Tables_levels_match.csv"  # noqa


def flatten(element, flat_list=None):
    '''Flatten given list'''
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        try:
            for item in element:
                flatten(item, flat_list)
        except TypeError:
            flat_list.append(element)
    else:
        flat_list.append(element)
    return flat_list


class Table_record():
    """ Object represent one connection between element ID
    and construction description. Description, ID and Vlaue
    description is analysing and return set of 6 parameters
    for every Table_record.
    Itis not revit element. Several Table_records could have the same
    element ID but different construction descriptions and parameters."""

    def __init__(self, constr_description, Id, value, match_dictionaries):
        self.constr_description = constr_description
        self.Id = Id
        self.value = value
        self.match_dictionaries = match_dictionaries
        self.level_1 = ""
        self.level_2 = ""
        self.level_3 = ""
        self.level_4 = ""
        self.level_5 = self.constr_description

    def set_parameters(self):
        """ Set Strabag parameters to records """
        for d in self.match_dictionaries:
            if d['Level 5'] == self.level_5:
                self.level_4 = d['Level 4']
                self.level_3 = d['Level 3']
                self.level_2 = d['Level 2']
                self.level_1 = d['Level 1']


class Revit_ID():
    """ Represent revit element Id with strabag parameters
    suitable_parameters - list with strings of
    suitable strabag parameters.
    Use this object to create Table_record object """
    def __init__(self, element_id, value, suitable_parameters):
        self.element_id = element_id
        self.value = value
        self.suitable_parameters = suitable_parameters

    def create_table_records(self, match_dictionaries):
        """ By suitable parameters generate Table_records """
        table_records = []
        for description in self.suitable_parameters:
            table_records.append(
                Table_record(
                    description,
                    self.element_id,
                    self.value,
                    match_dictionaries))
        return table_records


# class CSV_table(object):
#     """ CSV table """
#     def __init__(self, path, mode='r', test_diapason=[0, 10]):
#         self.path = path
#         self.mode = mode
#         self.test_diapason = test_diapason

#     @property
#     def strings(self):
#         '''Collect info from csv file.
#         last two symbols is '\n' '''
#         strings = []
#         with open(self.path, 'r') as f:
#             for x in f:
#                 last_symbol = str(x)[-1]
#                 if last_symbol == '\n':
#                     x = str(x)[:-1]
#                 strings.append(str(x))
#         return strings

#     @property
#     def titles(self):
#         """ Schedule titles
#         First string in schedule."""
#         titles = self.strings[0].split(",")
#         return titles

#     def size(self):
#         """ Return dictionary with schedule size """
#         return {
#             'Rows': len(self.records),
#             'Columns': len(self.titles)}

#     @property
#     def records(self):
#         """ Schedule records of elements """
#         if self.mode == 'r':
#             return self.strings[1:]
#         elif self.mode == 't':
#             return self.strings[
#                 (self.test_diapason[0] + 1):self.test_diapason[1]]


class Elements_schedule(CSV_table):
    """ CSV schedule from revit with element ID, value and Strabag parameters
    Columns: 1 - value, 2 - ID, 3... - Strabag descriptions
    Rows: 1 - titles, 2... - records for every Revit element
    mode t = testing, r = release"""

    def __init__(
            self, path, match_dictionaries, mode='r', test_diapason=[0, 10]):
        self.path = path
        self.mode = mode
        self.test_diapason = test_diapason
        self.match_dictionaries = match_dictionaries

    def create_Table_records(self):
        """ Create Table records from Revit_IDs objects """
        table_records = []
        for string in self.records:
            data = string.split(",")
            element_id = data[0]
            for i in range(1, len(self.titles)):
                if data[i] != "0" and data[i] != "":
                    table_records.append(
                        Table_record(
                            self.titles[i],
                            element_id,
                            data[i],
                            self.match_dictionaries))
        return table_records

    def set_Table_records_parameters(self):
        """ Set Table records match parameters."""

        table_records = self.create_Table_records()
        for record in table_records:
            record.set_parameters()
        return table_records

    def check_problems(self):
        """ If Id in problems list =>
        1 check if value of level_5 exist in match schedule
        2 check spaces in the end of titles
        3 check ',' in titles
        4 check match dictionaries, some spaces split one value to two
        If problems list is empty - all Table record elements
        find matched parameters in every levels"""

        table_records = self.set_Table_records_parameters()
        problems = []
        for record in table_records:
            if record.level_1 == '':
                problems.append(record.Id)
                problems.append(record.level_5)

        return False if len(problems) == 0 else problems

    def data_for_Power_BI(self):
        """ Return data to export in Excel """
        table_records = self.set_Table_records_parameters()
        excel_data = []
        if self.check_problems() is not False:
            return "Problems list:", self.check_problems()
        else:
            for table_record in table_records:
                record_data = []
                record_data.append(table_record.Id)
                record_data.append(table_record.value)
                record_data.append(table_record.level_1)
                record_data.append(table_record.level_2)
                record_data.append(table_record.level_3)
                record_data.append(table_record.level_4)
                record_data.append(table_record.level_5)
                excel_data.append(record_data)
        return excel_data


class Match_schedule(CSV_table):
    """ Schedule of match filter levels
    Rules = levels name: 'Level 1', 'Level 2' ...
            instead ',' - ' '(space)"""
    def __init__(self, path, mode='r', test_diapason=[0, 10]):
        self.path = path
        self.mode = mode
        self.test_diapason = test_diapason
        self.titles_r = self.titles[::-1]  # reverse titles
        self.levels_count = len(self.titles)
        self.dict_keys = [key for key in self.titles]
        self.dictionaries = self._create_dictionaries()

    def _create_dictionaries(self):
        """ Create dictionadies by last level and other levels """
        dicts = []
        for record in self.records:
            d = {}
            i = 0
            values = record.split(",")
            for key in self.dict_keys:
                d[key] = values[i]
                i += 1
            dicts.append(d)
        return dicts


""" Elements_schedule members """
match_schedule = Match_schedule(levels_match)
match_dictionaries = Match_schedule(levels_match).dictionaries
csv_table = CSV_table(levels_match)
schedule = Elements_schedule(
    file_ceilings, match_dictionaries, mode='r', test_diapason=[20, 30])
table_record = schedule.create_Table_records()[0]

print([
csv_table,
csv_table.strings,
# csv_table.titles,
# csv_table.records,
# csv_table.size(),

# schedule,
# schedule.strings,
# schedule.titles,
# schedule.records[20],
# schedule.records[22].split(","),
# schedule.size(),
# schedule._create_Revit_ID_elements(),
# schedule.create_Table_records(),
# schedule.set_Table_records_parameters(),
# schedule.check_problems(),
# schedule.data_for_Power_BI(),

# match_schedule,
# match_schedule.titles_r,
# match_schedule.records[0],
# match_schedule.dict_keys,
# match_schedule.dictionaries,
# match_schedule.size(),
# match_dictionaries,

# revit_id_element,
# revit_id_element.element_id,
# revit_id_element.value,
# revit_id_element.suitable_parameters,
# revit_id_element.create_table_records(match_dictionaries),

# table_record,
# table_record.constr_description,
# table_record.Id,
# table_record.value,
# table_record.set_parameters(),
# table_record.level_1,
# table_record.level_2,
# table_record.level_3,
# table_record.level_4,
# table_record.level_5,
# levels_match,
])

# OUT = schedule.data_for_Power_BI()
