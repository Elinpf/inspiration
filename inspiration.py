import xlrd
import xlwt
import xlutils.copy
import random


class Excel():

    def __init__(self, excel: str):
        self._excel_path = excel

        self._book = xlrd.open_workbook(excel, on_demand=True)
        self._book_write = None

        self._sheet: Sheet = None
        self._sheet_write = None

    def get_key_words(self):
        info = self.get_info()

        key_words = []
        for k, v in info.items():
            num = random.randint(0, v['length'] - 1)
            row = v['cell'][0]
            key_words.append(self._sheet.cell(row + num, v['cell'][1]).value)

        return key_words

    def get_info(self):
        deep = self.get_deep()

        data = {}
        for idx_col in range(1, self._sheet.ncols):
            data[idx_col] = {'composition': [],
                             'weight': 0, 'cell': [], 'length': 0}

            for idx_row in range(deep):

                # 单独提取weight
                if self._sheet.cell(idx_row, 0).value == 'weight':
                    data[idx_col]['weight'] = int(
                        self._sheet.cell(idx_row, idx_col).value)
                    continue

                # 单独提取length
                if self._sheet.cell(idx_row, 0).value == 'length':
                    data[idx_col]['length'] = int(
                        self._sheet.cell(idx_row, idx_col).value)
                    continue

                value = self._sheet.cell(idx_row, idx_col).value
                if value:
                    data[idx_col]['composition'].append(value)

            data[idx_col]['cell'] = [deep, idx_col]

        return data

    def select_sheet_by_name(self, name: str):
        self._sheet = self._book.sheet_by_name(name)

    def select_sheet(self, idx: int):
        self._sheet = self._book.sheet_by_index(idx)
        if self._book_write:
            self._sheet_write = self._book_write.get_sheet(idx)

    def get_deep(self):
        """
        级别的深度
        """
        deep = 0
        while True:
            row = self._sheet.row_values(deep)
            if not row[0]:
                break
            deep += 1
        return deep

    def upgrade(self):
        self._book_write = xlutils.copy.copy(self._book)

        for i in range(len(self._book.sheets())):
            self.select_sheet(i)
            self.upgrade_length()

        self.write()
        self.reopen()

    def upgrade_length(self):
        info = self.get_info()
        for k, v in info.items():
            row, col = v['cell']
            col_list = self.remove_empty(self._sheet.col_values(col)[row:])
            length = len(col_list)
            self._sheet_write.write(row-1, col, label=str(length))

    def write(self):
        self._book.release_resources()
        self._book_write.save(self._excel_path)

    def remove_empty(self, l: list):
        while '' in l:
            l.remove('')
        return l

    def reopen(self):
        self.close()
        self._book = xlrd.open_workbook(self._excel_path, on_demand=True)
        self._book_write = None

        self._sheet: Sheet = None
        self._sheet_write = None

    def close(self):
        self._book.release_resources()
        if self._book_write:
            # self._book_write.release_resources()
            del(self._book_write)
        del(self._book)


if __name__ == '__main__':
    excel = Excel('黑盒.xls')
    excel.upgrade()
    excel.select_sheet_by_name('构图')
    key_word_1 = excel.get_key_words()

    excel.select_sheet_by_name('主体')
    key_word_2 = excel.get_key_words()
    print(key_word_1, key_word_2)
