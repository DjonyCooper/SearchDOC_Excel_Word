import re
import traceback
from PyQt5.QtWidgets import (QWidget, QApplication, QPushButton, QGridLayout, QMessageBox, QLineEdit,
                             QStyle, QFileDialog, QPlainTextEdit, QLabel, QComboBox)
from PyQt5.QtCore import Qt, QPoint, QRectF
from PyQt5.QtGui import QPaintEvent, QPainter, QFont, QTextDocument
import pandas as pd
import locale, datetime
import ctypes
import openpyxl
from openpyxl.styles import Font
import zipfile, xml.etree.ElementTree
myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
locale.setlocale(category=locale.LC_ALL, locale="Russian")


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setMinimumSize(370, 220)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowFlags(Qt.FramelessWindowHint)

        self.grid = QGridLayout(self)
        self.grid.addWidget(self.label(), 1, 0, 1, 3)
        self.grid.addWidget(self.le_search_info(), 2, 0, 1, 2)
        self.grid.addWidget(self.cb_head_excel(), 2, 2, 1, 1)
        self.grid.addWidget(self.le_end_file_to_find(), 3, 0, 1, 3)
        self.grid.addWidget(self.le_dir_to_save_new_file(), 4, 0, 1, 3)
        self.grid.addWidget(self.button_search(), 5, 0, 1, 2)
        self.grid.addWidget(self.button_close(), 5, 2, 1, 1)
        self.grid.addWidget(self.le_info(), 6, 0, 1, 3)
        self.setLayout(self.grid)
        self.press = False
        self.last_pos = QPoint(0, 0)
        
        self.text = []
        self.articles = []
        self.new_articles = []

    def mouseMoveEvent(self, event):
        if self.press:
            self.move(event.globalPos() - self.last_pos)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.press = True
        self.last_pos = event.pos()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.press = False

    def paintEvent(self, event: QPaintEvent):
        painter = QPainter(self)

        painter.setOpacity(0.7)
        painter.setBrush(Qt.white)
        painter.drawRect(self.rect())

        document = QTextDocument()
        rect2 = QRectF(0, 0, 370, 220)
        document.setTextWidth(rect2.width())
        document.setHtml("<br><center><font size = '10' color='dark'>SearchDOC</font></center>"
                         "<center><font size = '3' color='grey'>разработчик DJC</font></center>")
        document.drawContents(painter, rect2)

    def label(self):
        label = QLabel()
        label.setMinimumSize(0, 70)

        return label
    def le_search_info(self):
        self.le_search_i = QLineEdit()
        self.le_search_i.setFocusPolicy(Qt.ClickFocus)
        self.le_search_i.setMinimumSize(207, 30)
        self.le_search_i.setPlaceholderText("Что искать?")
        serch_info_icon = self.le_search_i.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView),
                                                 QLineEdit.TrailingPosition)
        serch_info_icon.triggered.connect(self.browse_in_info)
        self.le_search_i.setFont(QFont('Century Gothic', 9))
        self.le_search_i.setAlignment(Qt.AlignCenter)
        self.le_search_i.returnPressed.connect(self.check_user_info)

        return self.le_search_i
    def le_end_file_to_find(self):
        self.le_file = QLineEdit()
        self.le_file.setMinimumSize(207, 30)
        self.le_file.setFocusPolicy(Qt.ClickFocus)
        self.le_file.setPlaceholderText("Где искать?")
        serch_info_icon = self.le_file.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView),
                                                        QLineEdit.TrailingPosition)
        serch_info_icon.triggered.connect(self.browse_in_files)
        self.le_file.setFont(QFont('Century Gothic', 9))
        self.le_file.setAlignment(Qt.AlignCenter)
        self.le_file.returnPressed.connect(self.check_user_info)

        return self.le_file

    def le_dir_to_save_new_file(self):
        self.le_save_new_file = QLineEdit()
        self.le_save_new_file.setMinimumSize(207, 30)
        self.le_save_new_file.setFocusPolicy(Qt.ClickFocus)
        self.le_save_new_file.setPlaceholderText("Куда сохранять?")
        serch_info_icon = self.le_save_new_file.addAction(self.style().standardIcon(QStyle.SP_FileDialogNewFolder),
                                                        QLineEdit.TrailingPosition)
        serch_info_icon.triggered.connect(self.browse_out_file)
        self.le_save_new_file.setFont(QFont('Century Gothic', 9))
        self.le_save_new_file.setAlignment(Qt.AlignCenter)
        self.le_save_new_file.returnPressed.connect(self.check_user_info)

        return self.le_save_new_file

    def le_info(self):
        self.le_info = QPlainTextEdit()
        self.le_info.setFont(QFont('Century Gothic', 9))
        self.le_info.setMinimumSize(300, 0)
        self.le_info.setReadOnly(True)
        self.le_info.setDisabled(True)
        bar = self.le_info.verticalScrollBar()
        bar.setValue(bar.maximum())
        return self.le_info

    def cb_head_excel(self):
        self.comb_box_excel = QComboBox()
        self.comb_box_excel.setMinimumSize(150, 30)

        return self.comb_box_excel
    def button_search(self):
        self.b_accept = QPushButton("Найти • Enter")
        self.b_accept.setMinimumSize(100, 30)
        self.b_accept.setShortcut('Enter')
        self.b_accept.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        self.b_accept.setStyleSheet("""QPushButton:!hover  {background-color: rgba(0, 0, 0, 5);
                                                            outline: none;
                                                            background-position: center;}
                                        QPushButton:hover  {border : 1px solid green;
                                                            outline: none;
                                                            background-color: rgba(0, 0, 0, 5);
                                                            background-position: center;}                    
                                        QPushButton:pressed{border : 1px solid dark;
                                                            outline: none;
                                                            background-color: green;
                                                            background-position: center;}
                                    """)
        self.b_accept.clicked.connect(self.check_user_info)

        return self.b_accept
    def button_close(self):
        self.b_close = QPushButton("Выход • ESC")
        self.b_close.setMinimumSize(100, 30)
        self.b_close.setShortcut('Esc')
        self.b_close.setIcon(self.style().standardIcon(QStyle.SP_DialogCloseButton))
        self.b_close.setStyleSheet("""QPushButton:!hover {background-color: rgba(0, 0, 0, 5);
                                                          outline: none;
                                                          background-position: center;}
                                      QPushButton:hover  {border : 1px solid red;
                                                          outline: none;
                                                          background-color: rgba(0, 0, 0, 5);
                                                          background-position: center;}                    
                                      QPushButton:pressed{border : 1px solid dark;
                                                          outline: none;
                                                          background-color: red;
                                                          background-position: center;}
                                     """)
        self.b_close.clicked.connect(self.close_app)
        return self.b_close

    def browse_in_info(self):
        browse_files_book_ost = QFileDialog.getOpenFileName(self, 'Выберите файл, содержащий артикулы...',
                                                            '', 'xlsx files (*.xlsx)')
        self.le_search_i.setText(browse_files_book_ost[0])
        self.check_name_excel()

    def browse_in_files(self):
        browse_files_book_ost = QFileDialog.getOpenFileName(self, 'Выберите файл, в котором нужно выполнить поиск...', '', 'docx files (*.docx)')
        self.le_file.setText(browse_files_book_ost[0])

    def browse_out_file(self):
        browse_files_book_ost = QFileDialog.getExistingDirectory(self, 'Выберите папку, куда сохранить новый файл...')
        self.le_save_new_file.setText(browse_files_book_ost)
    def check_name_excel(self):
        if self.le_search_i.text() != '':
            self.func_head_excel(self.le_search_i.text())
        else:
            pass

    def check_user_info(self):
        if self.le_search_i.text() != '':
            if self.le_file.text() != '':
                if self.le_save_new_file.text() != '':
                    self.func_msg_in_plain('Процесс сравнения данных начался... Ожидайте...')
                    self.func_search_info()
                else:
                    self.showMessageBox('Упс...', 'Для поиска необходимо заполнить все поля!')
            else:
                self.showMessageBox('Упс...', 'Для поиска необходимо заполнить все поля!')
        else:
            self.showMessageBox('Упс...', 'Для поиска необходимо заполнить все поля!')

    def func_head_excel(self, excel_file):
        df = pd.read_excel(f'{excel_file}')
        head = list(df.columns)
        self.comb_box_excel.addItems(head)

    def func_ext_data_from_excel(self):
        try:
            self.articles = []
            df = pd.read_excel(f'{self.le_search_i.text()}')
            self.articles = (df[f'{self.comb_box_excel.currentText()}'].tolist())

            return 'ok'
        except:
            return Exception.args

    def func_ext_data_from_word(self, name_doc):
        self.text = []
        try:
            WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            PARA = WORD_NAMESPACE + 'p'
            TEXT = WORD_NAMESPACE + 't'
            TABLE = WORD_NAMESPACE + 'tbl'
            ROW = WORD_NAMESPACE + 'tr'
            CELL = WORD_NAMESPACE + 'tc'

            with zipfile.ZipFile(f"{name_doc}") as docx:
                tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))

            for table in tree.iter(TABLE):
                for row in table.iter(ROW):
                    for cell in row.iter(CELL):
                        self.text.append(''.join(node.text for node in cell.iter(TEXT)))

            return 'ok'
        except Exception:
            return traceback.print_exc()

    def func_search_info(self):
        export_articules = self.func_ext_data_from_excel()
        if export_articules == 'ok':
            export_text = self.func_ext_data_from_word(f'{self.le_file.text()}')
            if export_text == 'ok':
                self.new_articles = []
                for articul in self.articles:
                    search = [s for s in self.text if articul in s]
                    if search != [] and search[0] != '':
                        delimiters = ",:"
                        split_search = re.split("|".join(delimiters), search[0])
                        clear = [s.lstrip() for s in split_search]
                        if articul in clear:
                            self.new_articles.append(articul)
                        else:
                            pass
                self.func_gen_new_excel()
            else:
                print(f'Ошибка при экспорте текста из Word: {export_text}')
        else:
            print(f'Ошибка при экспорте артикулов из Excel: {export_articules}')

    def func_gen_new_excel(self):
        wb = openpyxl.Workbook()
        list = wb.active
        list.append([f'{self.comb_box_excel.currentText()}'])
        list.column_dimensions['A'].width = 15
        list['A1'].font = Font(color="FF0000", bold=True)
        for art in self.new_articles:
            list.append([art])
        wb.save(f'{self.le_save_new_file.text()}/File_{datetime.datetime.now().strftime("%d%m%y%H%M%S")}.xlsx')
        self.func_msg_in_plain(f'Совпадения найдены у {len(self.new_articles)} артикулов.\n'
                               f'Файл с совпадениями успешно создан и сохранен:\n'
                               f'{self.le_save_new_file.text()}/File_{datetime.datetime.now().strftime("%d%m%y%H%M%S")}.xlsx')

    def func_msg_in_plain(self, text):
        now_time_date = datetime.datetime.now()
        self.le_info.insertPlainText(f'\n{now_time_date.strftime("%H:%M:%S")} - {text}')
        bar = self.le_info.verticalScrollBar()
        bar.setValue(bar.maximum())

    def showMessageBox(self, title, message):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(self.style().standardIcon(QStyle.SP_BrowserStop))
        msgBox.setWindowTitle(title)
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText(message)
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec_()

    def close_app(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(self.style().standardIcon(QStyle.SP_TitleBarContextHelpButton))
        msgBox.setWindowTitle("Выйти")
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText("Вы уверены, что хотите выйти?")
        msgBox.setStyleSheet("font: 12px;"
                             "font-family: Century Gothic;")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        _msgBoxClose = msgBox.exec_()
        if _msgBoxClose == QMessageBox.Yes:
            self.close()

if __name__ == ('__main__'):
    import sys
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())