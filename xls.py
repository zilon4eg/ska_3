import xlwings as xw
import datetime


class Excel:
    def __init__(self, xls_path=r'C:\Users\suhorukov.iv\Desktop\test.xlsx'):
        self.xls_path = xls_path
        self.app = xw.App(visible=False)
        self.wb = self.app.books.open(file_path)

    def save_close_quit(self):
        self.wb.save(self.xls_path)
        self.wb.close()
        self.app.quit()

    def work_sheet(self):
        ws_name1 = datetime.datetime.today().strftime('%d-%m-%Y')
        ws_name2 = datetime.datetime.today().strftime('%d-%m-%Y %H-%M-%S')
        return ws_name1 if ws_name1 not in list(sheet.name for sheet in self.wb.sheets) else ws_name2


if __name__ == '__main__':
    file_name = r'test.xlsx'
    file_path = f'C:\\Users\\suhorukov.iv\\Desktop\\{file_name}'
    # print(datetime.datetime.today().strftime('%d-%m-%Y'))
    # print(datetime.datetime.today().strftime('%d-%m-%Y %H-%M-%S'))
    # app = xlwings.App(file_path).visible = False
    # wb = xw.Book()  # this will create a new workbook

    # wb = xw.App(file_path).visible = False
    # wb = xw.Book(file_path)  # on Windows: use raw strings to escape backslashes
    # sheet = wb.sheets[0] #0,1,2... or 'name'
    # print(sheet.range('A1').value)

    # app = xw.App(visible=False)
    # wb = app.books.open(file_path)
    # print(list(sheet.name for sheet in wb.sheets))
    # ws = wb.sheets[0]
    # ws.range('B2').value = '13413434134'
    # wb.save(file_path)
    # wb.close()
    # app.quit()


    # xw.Book.save()
    # xw.Book.close()
    # wb.save()
    # wb.close()
    # app.quit()
    xxl = Excel(file_path)
    print(xxl.work_sheet())
