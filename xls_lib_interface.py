import xlwings as xw

# Encapsula funções de interface com a biblioteca base de mapipulação de arquivos XLS.
class Xls_lib_interface:
    def __init__(self, nome_xls):
        self.wb = xw.Book(nome_xls)
        self.ws = self.wb.sheets[0]
    
    def sheets(self, index):
        self.ws = self.wb.sheets[index]
    
    def edit_val_cell(self, cell, val):
        self.ws.range(cell).value = val
    
    def edit_font_size_cell(self, cell, size):
        self.ws.range(cell).api.Font.Size = size
