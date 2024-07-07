from openpyxl import load_workbook
from pprint import pprint


class EquiqXlsx:

    def __init__(self, filename):
        self.filename = filename

        self.col_id = {
            "Descr. Sint.": None,
            "Dt.Aquisicao": None,
            "Quantidade": None,
            "Tipo Ativo": None,
        }
        self.data = []

    def load(self):
        self.wb = load_workbook(filename=self.filename)
        self.ws = self.wb.active

    def get_col_id(self, row):
        for cell in row:
            if cell.value in self.col_id:
                self.col_id[cell.value] = cell.column_letter
        pprint(self.col_id)

    def data_append_item(self, idx, row):
        item = {}
        for name, letter in self.col_id.items():
            item[name] = self.ws[f'{letter}{idx}'].value
        self.data.append(item)

    def walk_through(self):
        for idx_row, row in enumerate(self.ws.iter_rows(max_row=5), start=1):
            if idx_row == 1:
                self.get_col_id(row)
            else:
                self.data_append_item(idx_row, row)
    
    def print(self):
        pprint(self.data)

    def process(self):
        self.load()
        self.walk_through()
        self.print()
    


if __name__ == '__main__':
    EquiqXlsx('equipamentos.xlsx').process()
