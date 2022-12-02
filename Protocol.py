import sys
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6 import uic
from PyQt6.QtWidgets import QFileDialog, QDialog
from docx import Document
from docxtpl import DocxTemplate



class MainWindow(QtWidgets.QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ui = uic.loadUi("mydesign.ui", self)
        self.ui.pushButton.clicked.connect(self.create_file)

    def open_file_dialog(self):
        dialog = QFileDialog(self)
        dialog.setDirectory(r'Z:\Для всех\НТОИ\Вертаев\2022\11 Ноябрь')
        dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dialog.setNameFilter("Word (*.doc *.docx)")
        dialog.setViewMode(QFileDialog.ViewMode.List)
        if dialog.exec():
            filenames = dialog.selectedFiles()
        return filenames

    def get_context(self):
        context_lst = {}
        context_lst['pz_number'] = self.ui.pz_number_line.text()
        context_lst['definition'] = self.correct_strip(self.ui.name_line.text())
        context_lst['acronym']=self.ui.acronymComboBox.currentText()
        context_lst['producer'] = self.ui.producerComboBox.currentText()
        context_lst['model'] = self.ui.model_line.text()
        context_lst['protocol_number'] = self.ui.protocol_number_line.text()
        context_lst['letter'] = self.ui.letter_line.text()
        context_lst['sample_1'] = self.ui.sample_1.text()
        context_lst['sample_2'] = self.ui.sample_2.text()
        context_lst['sample_3'] = self.ui.sample_3.text()
        context_lst['num_sample1'] = self.ui.sample_1_num.text()
        context_lst['num_sample2'] = self.ui.sample_2_num.text()
        context_lst['num_sample3'] = self.ui.sample_3_num.text()
        context_lst['metrolog'] = self.ui.metrologComboBox.currentText()
        context_lst['tester'] = self.ui.testerComboBox.currentText()
        print(context_lst)
        return context_lst

    def create_file(self):
        '''Процесс создания файла '''
        files=self.open_file_dialog()
        for file in files:
            self.insert_setup(file)
            self.update_fields(file)


    def insert_setup(self, path):
        '''Вставляем ссылку setup в конец документа'''
        doc = Document(path)
        doc.add_paragraph('{{ setup }}')
        doc.save(path)


    def update_fields(self, file):
        '''Вставляем поля из файла fields и затем заполняем их'''
        doc = DocxTemplate(file)
        sd = doc.new_subdoc('fields.docx')
        context = {'setup': sd}
        doc.render(context)
        doc.save(file)
        context = self.get_context()
        doc = DocxTemplate(file)
        doc.render(context) # Падает тут
        doc.save(file)
        print('Выполнено')

    @staticmethod
    def correct_strip(text):
        '''Херня полная, но пока впадлу кавычки обрабатывать'''
        text=text.replace('«','')
        text=text.replace('»', '')
        return text

app = QtWidgets.QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
