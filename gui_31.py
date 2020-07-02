# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gui_v3.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from main_logic import gsheets, doc_parts
from classes import pm_doc, specie, site
import logging



#Creating the logger class
class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QtWidgets.QPlainTextEdit(parent)
        self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)




class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1185, 910)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.lbl_info = QtWidgets.QLabel(self.centralwidget)
        self.lbl_info.setGeometry(QtCore.QRect(10, 10, 1171, 141))
        self.lbl_info.setFrameShape(QtWidgets.QFrame.Box)
        self.lbl_info.setObjectName("lbl_info")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 220, 491, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setUnderline(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(700, 170, 20, 631))
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")

        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setGeometry(QtCore.QRect(710, 690, 471, 20))
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")

        self.tb_xls = QtWidgets.QToolButton(self.centralwidget)
        self.tb_xls.setGeometry(QtCore.QRect(720, 160, 451, 41))
        self.tb_xls.setObjectName("tb_save")

        self.tb_save = QtWidgets.QToolButton(self.centralwidget)
        self.tb_save.setGeometry(QtCore.QRect(720, 210, 451, 41))
        self.tb_save.setObjectName("tb_save")

        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(10, 770, 451, 51))
        self.checkBox.setObjectName("checkBox")

        self.tb_img = QtWidgets.QToolButton(self.centralwidget)
        self.tb_img.setGeometry(QtCore.QRect(720, 260, 451, 41))
        self.tb_img.setObjectName("tb_img")

        self.tb_maps = QtWidgets.QToolButton(self.centralwidget)
        self.tb_maps.setGeometry(QtCore.QRect(720, 310, 451, 41))
        self.tb_maps.setObjectName("tb_maps")

        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(20, 820, 1161, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")

        self.bt_main = QtWidgets.QPushButton(self.centralwidget)
        self.bt_main.setGeometry(QtCore.QRect(720, 710, 331, 91))
        font = QtGui.QFont()
        font.setPointSize(22)
        font.setBold(True)
        font.setWeight(75)
        self.bt_main.setFont(font)
        self.bt_main.setObjectName("bt_main")

        self.layoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.layoutWidget.setGeometry(QtCore.QRect(10, 160, 481, 41))
        self.layoutWidget.setObjectName("layoutWidget")

        self.site = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.site.setContentsMargins(0, 0, 0, 0)
        self.site.setObjectName("site")

        self.lbl_site = QtWidgets.QLabel(self.layoutWidget)
        self.lbl_site.setObjectName("lbl_site")
        self.site.addWidget(self.lbl_site)

        self.cb_site = QtWidgets.QComboBox(self.layoutWidget)
        self.cb_site.setObjectName("cb_site")
        self.site.addWidget(self.cb_site)


        self.bt_cons = QtWidgets.QPushButton(self.centralwidget)
        self.bt_cons.setGeometry(QtCore.QRect(1070, 710, 101, 91))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.bt_cons.setFont(font)
        self.bt_cons.setObjectName("bt_cons")

        self.base_lst = QtWidgets.QListWidget(self.centralwidget)
        self.base_lst.setGeometry(QtCore.QRect(10, 250, 300, 521))
        self.base_lst.setObjectName("base_lst")
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.base_lst.addItem(item)



        self.doc_lst = QtWidgets.QListWidget(self.centralwidget)
        self.doc_lst.setGeometry(QtCore.QRect(400, 250, 300, 521))
        self.doc_lst.setObjectName("doc_lst")

        self.bt_allr = QtWidgets.QPushButton(self.centralwidget)
        self.bt_allr.setGeometry(QtCore.QRect(320, 260, 71, 23))
        self.bt_allr.setObjectName("bt_allr")

        self.bt_alll = QtWidgets.QPushButton(self.centralwidget)
        self.bt_alll.setGeometry(QtCore.QRect(320, 360, 71, 23))
        self.bt_alll.setObjectName("bt_alll")

        self.bt_oner = QtWidgets.QPushButton(self.centralwidget)
        self.bt_oner.setGeometry(QtCore.QRect(320, 290, 71, 23))
        self.bt_oner.setObjectName("bt_oner")

        self.bt_onel = QtWidgets.QPushButton(self.centralwidget)
        self.bt_onel.setGeometry(QtCore.QRect(320, 330, 71, 23))
        self.bt_onel.setObjectName("bt_onel")

        self.bt_sph = QtWidgets.QPushButton(self.centralwidget)
        self.bt_sph.setGeometry(QtCore.QRect(320, 450, 71, 23))
        self.bt_sph.setObjectName("bt_sph")

        self.bt_up = QtWidgets.QPushButton(self.centralwidget)
        self.bt_up.setGeometry(QtCore.QRect(460, 780, 81, 23))
        self.bt_up.setObjectName("bt_up")

        self.bt_dw = QtWidgets.QPushButton(self.centralwidget)
        self.bt_dw.setGeometry(QtCore.QRect(570, 780, 81, 23))
        self.bt_dw.setObjectName("bt_dw")

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(720, 360, 111, 16))
        self.label_3.setObjectName("label_3")

        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(1060, 850, 121, 16))
        self.label_4.setObjectName("label_4")

        MainWindow.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1185, 21))
        self.menubar.setObjectName("menubar")
        self.menuSettings = QtWidgets.QMenu(self.menubar)
        self.menuSettings.setObjectName("menuSettings")
        MainWindow.setMenuBar(self.menubar)

        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionAdd_site = QtWidgets.QAction(MainWindow)
        self.actionAdd_site.setObjectName("actionAdd_site")
        self.menuSettings.addAction(self.actionAdd_site)
        self.menubar.addAction(self.menuSettings.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


















#My code in Ui_MainWindow

    #Setting the buttons connections
        #Setting the sites in site combo box
        for i in range(1,172):
            if len(str(i)) == 1:
                i = 'ROSPA000'+str(i)
            elif len(str(i)) == 2:
                i = 'ROSPA00'+str(i)
            elif len(str(i)) == 3:
                i = 'ROSPA0'+str(i)
            self.cb_site.addItem(i)

        #Setting the select folders buttons
        self.tb_save.clicked.connect(lambda: self.fp_dialog(self.tb_save))
        self.tb_img.clicked.connect(lambda: self.fp_dialog(self.tb_img))
        self.tb_maps.clicked.connect(lambda: self.fp_dialog(self.tb_maps))
        self.tb_xls.clicked.connect(lambda: self.fn_dialog(self.tb_xls))


        #Setting the list buttons
        self.bt_alll.clicked.connect(self.move_all_items_left)
        self.bt_allr.clicked.connect(self.move_all_items_right)
        self.bt_onel.clicked.connect(self.move_one_item_left)
        self.bt_oner.clicked.connect(self.move_one_item_right)
        self.bt_sph.clicked.connect(self.add_sp_header)

        self.bt_up.clicked.connect(self. move_up)
        self.bt_dw.clicked.connect(self.move_dw)

        self.bt_cons.clicked.connect(self.test_func)

        self.bt_main.clicked.connect(self.create_doc_func)

        #Setting the log box
        self.lgbx = QTextEditLogger(self.centralwidget)
        self.lgbx.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.lgbx)
        logging.getLogger().setLevel(logging.DEBUG)
        self.vLayout = QtWidgets.QWidget(self.centralwidget)
        self.vLayout.setGeometry(QtCore.QRect(720, 380, 451, 311))
        self.vl = QtWidgets.QVBoxLayout(self.vLayout)
        self.vl.addWidget(self.lgbx.widget)



#Setting diferent functions in the main class
    #The function that opens the select folders buttons
    def fp_dialog(self,btn):
        try:
            self.foldername = QFileDialog.getExistingDirectory()
            if self.foldername != '':
                btn.setText(self.foldername)
                btn.setStyleSheet("background-color: #98FB98")
        except Exception as err:
            logging.debug(err)

    #The function that opens the select xls dialog
    def fn_dialog(self,btn):
        try:
            self.fn = QFileDialog.getOpenFileName()[0]
            if self.fn != '':
                btn.setText(self.fn)
                btn.setStyleSheet("background-color: #98FB98") 
        except Exception as err:
            logging.debug(err)


    #The functions for list buttons
    def move_one_item_left(self):
        try:
            self.base_lst.addItem(self.doc_lst.takeItem(self.doc_lst.currentRow()))
        except Exception as err:
            logging.debug(err)
    def move_one_item_right(self):
        try:
            self.doc_lst.addItem(self.base_lst.takeItem(self.base_lst.currentRow()))
        except Exception as err:
            logging.debug(err)
    def move_all_items_right(self):
        try:
            while self.base_lst.count() > 0:
                self.doc_lst.addItem(self.base_lst.takeItem(0))
        except Exception as err:
            logging.debug(err)
    def move_all_items_left(self):
        try:
            while self.doc_lst.count() > 0:
                self.base_lst.addItem(self.doc_lst.takeItem(0))
        except Exception as err:
            logging.debug(err)
    def add_sp_header(self):
        try:
            self.doc_lst.addItem('Start species iterations')
        except Exception as err:
            logging.debug(err)

    def move_up(self):
        try:
            row = self.doc_lst.currentRow()
            currentItem = self.doc_lst.takeItem(row)
            self.doc_lst.insertItem(row - 1, currentItem)
            self.doc_lst.setCurrentRow(row - 1)
        except Exception as err:
            logging.debug(err)

    def move_dw(self):
        try:
            row = self.doc_lst.currentRow()
            currentItem = self.doc_lst.takeItem(row)
            self.doc_lst.insertItem(row + 1, currentItem)
            self.doc_lst.setCurrentRow(row + 1)
        except Exception as err:
            logging.debug(err)

    #Test function    
    def test_func(self):
        try:
            logging.debug(self.fn)
        except Exception as err:
            logging.debug(err)



    #The main function responsible for result
    def create_doc_func(self):
        try:
            #Getting the parameters
            self.save_path = str(self.tb_save.text())
            logging.info('Save path: {}'.format(self.save_path))

            self.img_path = str(self.tb_img.text())
            logging.info('Img path: {}'.format(self.img_path))

            self.maps_path = str(self.tb_maps.text())
            logging.info('Maps path: {}'.format(self.maps_path))

            self.xls_path = str(self.tb_xls.text())
            logging.info('Xls path: {}'.format(self.xls_path))

            self.empty_tables = not self.checkBox.isChecked()
            logging.info('Empty tables: {}'.format(self.empty_tables))

            self.doc_format = tuple([self.doc_lst.item(x).text() for x in range(self.doc_lst.count())])
            logging.info('Doc format: {}'.format(self.doc_format))

            #Creating the document
            self.doc = pm_doc()
            logging.debug('The document object was initialised!')

            #Creating the site tables
            self.site = site(str(self.cb_site.currentText()))
            logging.debug('The site object was created!')
            self.bf = self.site.bf()
            logging.debug('The info site dataframe was created')
            self.master = self.site.master(self.xls_path)
            logging.debug('The master dataframe was created')
            self.impacts = self.site.impacts(self.xls_path)
            logging.debug('The impacts dataframe was created')
            self.masuri = self.site.masuri(self.xls_path)
            logging.debug('The measures dataframe was created')
            self.descrieri = self.site.descrieri(self.xls_path)
            logging.debug('The descriptions dataframe was created')
            
            #The list of species
            self.df_rows = list(self.master.index)
            logging.debug('This is the species list: {}'.format(self.df_rows))

            #This should create the doc_format list of lists based and without the Start species iterations 
            self.sphs = [x for x,y in enumerate(self.doc_format) if y == 'Start species iterations']
            self.sphs.append(len(self.doc_format))
            self.doc_chapters = [self.doc_format[x+1:y] for x, y in zip(self.sphs, self.sphs[1:])]
            logging.debug('This are the chapters: {}'.format(self.doc_chapters))


            #The main iteration. It suppose to create the document
            for chapter in self.doc_chapters:
                logging.info('Started working on ' + ' '.join(chapter))
                for row in self.df_rows:
                    #Init the species object               
                    self.sp = specie(row, self.master, self.bf, self.descrieri, self.impacts, self.masuri)
                    logging.info('Created the object for ' + self.sp.lat_sp)
                    self.doc.sp_header(self.sp)
                    for part in chapter:
                        #Create a function from the string title of the part and run it
                        part_func = getattr(self.doc, doc_parts[part])
                        part_func(self.empty_tables, self.sp)

            self.doc.save(str(self.cb_site.currentText()) +'_export'+'.docx')
        except Exception as err:
            logging.debug('A crapat!')
            logging.debug(err)























    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.lbl_info.setText(_translate("MainWindow", "Here is the informative text"))
        self.label.setText(_translate("MainWindow", "Doc format! (Choose the chapters/tables in the desired order)"))
        self.tb_save.setText(_translate("MainWindow", "Choose where to save the file!"))
        self.checkBox.setText(_translate("MainWindow", "I want empty tables! If checked the tables will be exported empty."))
        self.tb_img.setText(_translate("MainWindow", "Choose the species photos folder!"))
        self.tb_maps.setText(_translate("MainWindow", "Choose the species maps folder!"))
        self.tb_xls.setText(_translate("MainWindow", "Choose the excel file!"))
        self.bt_main.setText(_translate("MainWindow", "Create document!"))
        self.lbl_site.setText(_translate("MainWindow", "Select the site!"))
        self.bt_cons.setText(_translate("MainWindow", "Conservation"))
        __sortingEnabled = self.base_lst.isSortingEnabled()
        self.base_lst.setSortingEnabled(False)
        item = self.base_lst.item(0)
        item.setText(_translate("MainWindow", "Descriptive table A"))
        item = self.base_lst.item(1)
        item.setText(_translate("MainWindow", "Descriptive table B"))
        item = self.base_lst.item(2)
        item.setText(_translate("MainWindow", "Conservation table A"))
        item = self.base_lst.item(3)
        item.setText(_translate("MainWindow", "Matrix 1 table"))
        item = self.base_lst.item(4)
        item.setText(_translate("MainWindow", "Conservation table B"))
        item = self.base_lst.item(5)
        item.setText(_translate("MainWindow", "Matrix 2 table"))
        item = self.base_lst.item(6)
        item.setText(_translate("MainWindow", "Matrix 3 table"))
        item = self.base_lst.item(7)
        item.setText(_translate("MainWindow", "Conservation table C"))
        item = self.base_lst.item(8)
        item.setText(_translate("MainWindow", "Matrix 4 table"))
        item = self.base_lst.item(9)
        item.setText(_translate("MainWindow", "Matrix 5 table"))
        item = self.base_lst.item(10)
        item.setText(_translate("MainWindow", "Matrix 6 table"))
        item = self.base_lst.item(11)
        item.setText(_translate("MainWindow", "Conservation table D"))
        item = self.base_lst.item(12)
        item.setText(_translate("MainWindow", "Matrix 7 table"))
        item = self.base_lst.item(13)
        item.setText(_translate("MainWindow", "Measures chapter"))





        self.base_lst.setSortingEnabled(__sortingEnabled)
        self.bt_allr.setText(_translate("MainWindow", ">>"))
        self.bt_alll.setText(_translate("MainWindow", "<<"))
        self.bt_oner.setText(_translate("MainWindow", ">"))
        self.bt_onel.setText(_translate("MainWindow", "<"))
        self.bt_sph.setText(_translate("MainWindow", "S I"))
        self.bt_up.setText(_translate("MainWindow", "Up"))
        self.bt_dw.setText(_translate("MainWindow", "Down"))
        self.label_3.setText(_translate("MainWindow", "Log window"))
        self.label_4.setText(_translate("MainWindow", "@ Dorin Damoc - 2020"))
        self.menuSettings.setTitle(_translate("MainWindow", "Settings"))
        self.actionAdd_site.setText(_translate("MainWindow", "Add site"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
