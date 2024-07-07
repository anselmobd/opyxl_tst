from openpyxl import load_workbook
from pprint import pprint


def main():
    wb = load_workbook(filename = 'equipamentos.xlsx')
    ws = wb.active

    col_name = {
        "Descr. Sint.": None,
        "Dt.Aquisicao": None,
        "Quantidade": None,
        "Tipo Ativo": None,
    }

    for idx_row, row in enumerate(ws.iter_rows(max_row=5), start=1):
        if idx_row == 1:
            for cell in row:
                if cell.value in col_name:
                    col_name[cell.value] = cell.column_letter
            pprint(col_name)


if __name__ == '__main__':
    main()