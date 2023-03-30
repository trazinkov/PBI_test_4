# -*- coding: utf-8 -*-

class CSV_table(object):
    """ CSV table """
    def __init__(self, path, mode='r', test_diapason=[0, 10]):
        self.path = path
        self.mode = mode
        self.test_diapason = test_diapason

    @property
    def strings(self):
        '''Collect info from csv file.
        last two symbols is '\n' '''
        strings = []
        with open(self.path, 'r') as f:
            for x in f:
                last_symbol = str(x)[-1]
                if last_symbol == '\n':
                    x = str(x)[:-1]
                strings.append(str(x))
        return strings

    @property
    def titles(self):
        """ Schedule titles
        First string in schedule."""
        titles = self.strings[0].split(",")
        return titles

    def size(self):
        """ Return dictionary with schedule size """
        return {
            'Rows': len(self.records),
            'Columns': len(self.titles)}

    @property
    def records(self):
        """ Schedule records of elements """
        if self.mode == 'r':
            return self.strings[1:]
        elif self.mode == 't':
            return self.strings[
                (self.test_diapason[0] + 1):self.test_diapason[1]]
