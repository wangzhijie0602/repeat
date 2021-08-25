from subprocess import run
import sys,Ui_FormUi,threading,Backend
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow


def get_file(filelabel,sheetlist,rangecombom):
    name = QFileDialog.getOpenFileName(None,"选择文件", "/", "xlsx files (*.xlsx);;xls files (*.xls);;all files (*)")
    if name[0] != "":
        path = name[0]
        filelabel.setText(path)
        sheet_list = Backend.GetSheet(path)
        sheetlist.clear()
        for i in sheet_list:
            sheetlist.addItem(i)
        sheetlist.setCurrentRow(0)
        rangecombom.setEnabled(True)
        if ui.rangecombomP.isEnabled() and ui.rangecombomT.isEnabled():
            ui.runbutton.setEnabled(True)


def run():
    path = ui.filelabelT.text()
    if path[-1].upper() == "X":
        row = "1048576"
    else:
        row = "65536"
    list = Backend.DateRead(path,ui.sheetlistT.currentItem().text(),ui.rangecombomT.itemText(ui.rangecombomT.currentIndex())[0],row)
    list = filter(None,list)
    t = threading.Thread(target=Backend.run,args=(ui.filelabelP.text(),ui.sheetlistP.currentItem().text(),ui.rangecombomP.itemText(ui.rangecombomP.currentIndex())[0],list,),daemon = True)
    t.start()


def main():
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    global ui
    ui = Ui_FormUi.Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.getfilebuttonP.clicked.connect(lambda:get_file(ui.filelabelP,ui.sheetlistP,ui.rangecombomP))
    ui.getfilebuttonT.clicked.connect(lambda:get_file(ui.filelabelT,ui.sheetlistT,ui.rangecombomT))
    ui.runbutton.clicked.connect(run)
    ui.exitbutton.clicked.connect(sys.exit)
    ui.runbutton.setEnabled(False)
    ui.rangecombomP.setEnabled(False)
    ui.rangecombomT.setEnabled(False)
    for asc in range(65,90 + 1):
        ui.rangecombomP.addItem("{}列".format(chr(asc)))
        ui.rangecombomT.addItem("{}列".format(chr(asc)))
    MainWindow.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()