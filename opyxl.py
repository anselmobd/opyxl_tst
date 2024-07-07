from openpyxl import load_workbook
from pprint import pprint


class EquiqXlsx:

    def __init__(self, filename):
        self.filename = filename

        self.col_name = {
            "Descr. Sint.": None,
            "Dt.Aquisicao": None,
            "Quantidade": None,
            "Tipo Ativo": None,
        }
        self.data = []

    def process(self):
        self.wb = load_workbook(filename=self.filename)
        self.ws = self.wb.active

        for idx_row, row in enumerate(self.ws.iter_rows(max_row=5), start=1):
            if idx_row == 1:
                for cell in row:
                    if cell.value in self.col_name:
                        self.col_name[cell.value] = cell.column_letter
                pprint(self.col_name)
            else:
                data_row = {}
                for name, letter in self.col_name.items():
                    data_row[name] = self.ws[f'{letter}{idx_row}'].value
                self.data.append(data_row)
        pprint(self.data)
    


if __name__ == '__main__':
    EquiqXlsx('equipamentos.xlsx').process()
