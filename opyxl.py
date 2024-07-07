from openpyxl import load_workbook
from pprint import pprint


class EquiqXlsx:

    def __init__(self, filename):
        self.filename = filename

        self.columns_idx = {
            'Descr. Sint.': None,
            'Dt.Aquisicao': None,
            'Quantidade': None,
            'Tipo Ativo': None,
        }
        self.data = []

    def load(self):
        self.wb = load_workbook(filename=self.filename)
        self.ws = self.wb.active

    def get_col_id(self, row):
        for idx_col, cell in enumerate(row):
            if cell.value in self.columns_idx:
                self.columns_idx[cell.value] = idx_col
        pprint(self.columns_idx)

    def data_append_item(self, idx_row, row):
        item = {}
        for name, idx_col in self.columns_idx.items():
            item[name] = row[idx_col].value
        self.data.append(item)

    def item_valido(self, row):
        return not row[self.columns_idx['Tipo Ativo']].value

    def walk_through(self):
        for idx_row, row in enumerate(self.ws.iter_rows(max_row=5), start=1):
            if idx_row == 1:
                self.get_col_id(row)
            else:
                if self.item_valido(row):
                    self.data_append_item(idx_row, row)

    def print(self):
        pprint(self.data)

    def process(self):
        self.load()
        self.walk_through()
        self.print()


if __name__ == '__main__':
    EquiqXlsx('equipamentos.xlsx').process()
