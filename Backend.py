from Mian import main
import xlwings as xw
import time


def GetSheet(bookname):
    app = xw.App(visible=False,add_book=False)
    wb = app.books.open(bookname)
    list = []
    num = len(wb.sheets)
    for i in range(0,num):
        sht = wb.sheets[i]
        list.append(sht.name)
    wb.close()
    app.kill()
    return list


def input_error(sht):
    sht.color = 255,0,0


def DateRead(book,sheet,column,row):
    app = xw.App(visible=False,add_book=False)
    wb = app.books.open(book)
    sht = wb.sheets[sheet]
    StartRow = column + "1"
    EndRow = column + row
    RangeDate = sht.range(StartRow + ":" + EndRow).value
    wb.close()
    app.kill()
    return RangeDate
    

def run(book,sheet,column,list):
    app = xw.App(visible=True,add_book=False)
    wb = app.books.open(book)
    sht = wb.sheets[sheet]
    num = 0
    none = 0
    while True:
        num += 1
        table = sht.range(column + str(num))
        if table.value == None or table.value == "":
            none += 1
            if none != 10:
                continue
            else:
                break
        else:
            none = 0
        table.color = 255,255,0
        time.sleep(0.01)
        if table.value not in list:
            input_error(table)
    

def main():
    return


if __name__ == "__main__":
    main()