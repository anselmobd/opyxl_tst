from openpyxl import load_workbook
from pprint import pprint


class EquiqXlsx:

    # columns read
    DESCR = 'Descr. Sint.'
    DATA = 'Dt.Aquisicao'
    QUANT = 'Quantidade'
    TIPO = 'Tipo Ativo'

    def __init__(self, filename):
        self.filename = filename

        self.origin_columns_idx = {
            self.DESCR: None,
            self.DATA: None,
            self.QUANT: None,
            self.TIPO: None,
        }
        self.origin_data = []
        self.destination = {}
        self.info_recipe = {
            'key':(
                {
                    'value': self.DESCR,
                    'transform': self.get_tag_descricao
                },
                {
                    'value': self.DATA,
                    'transform': self.get_ano_data
                },
            ),
            'value': {
                'value': self.QUANT,
            },
            'apply': self.aualiza_valor,
        }
        self.tags = {
            'MATRICIAL': 'matricial',
            'SERVIDOR': 'servidor',
            'MICROCOMPUTADOR': 'desktop',
            'PROJETOR': 'projetor',
            'MONITOR': 'monitor',
            None: 'outros',
        }
        self.max_row=25

    def get_tag_descricao(self, value):
        for search, tag in self.tags.items():
            if search and value.find(search) > -1:
                return tag
        return self.tags[None]

    def get_ano_data(self, value):
        return value.year

    def aualiza_valor(self, key, value):
        try:
            old_value = self.destination[key]
        except KeyError as _:
            old_value = 0
        self.destination[key] = old_value + value

    def mount_info(self, recipe, row):
        if isinstance(recipe, tuple):
            step_info = []
            for step in recipe:
                step_info.append(self.mount_info(step, row))
            return tuple(step_info)
        else:
            value = row[recipe['value']]
            if 'transform' in recipe:
                value = recipe['transform'](value)
            return value

    def mount_destination(self):
        for row in self.origin_data:
            key = self.mount_info(self.info_recipe['key'], row)
            value = self.mount_info(self.info_recipe['value'], row)
            self.info_recipe['apply'](key, value)
        return 1

    def load(self):
        self.wb = load_workbook(filename=self.filename)
        self.ws = self.wb.active

    def get_col_id(self, row):
        for idx_col, cell in enumerate(row):
            if cell.value in self.origin_columns_idx:
                self.origin_columns_idx[cell.value] = idx_col
        pprint(self.origin_columns_idx)

    def data_append_item(self, idx_row, row):
        item = {}
        for name, idx_col in self.origin_columns_idx.items():
            item[name] = row[idx_col].value
        self.origin_data.append(item)

    def item_valido(self, row):
        return not row[self.origin_columns_idx['Tipo Ativo']].value

    def walk_through(self):
        for idx_row, row in enumerate(
                self.ws.iter_rows(max_row=self.max_row), start=1):
            if idx_row == 1:
                self.get_col_id(row)
            else:
                if self.item_valido(row):
                    self.data_append_item(idx_row, row)

    def print(self):
        pprint(self.origin_data)
        pprint(self.destination)

    def process(self):
        self.load()
        self.walk_through()
        self.mount_destination()
        self.print()


if __name__ == '__main__':
    EquiqXlsx('equipamentos.xlsx').process()
