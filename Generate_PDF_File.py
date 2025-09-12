import os
import pandas
import PyPDF2
from Window import Ui_GUI
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys

class ShowUI(QMainWindow, Ui_GUI):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.Remove_Students_list = []
        self.SelectExcelPath_button.clicked.connect(self.__OpenFileDialog)
        self.SelectPDFPath_button.clicked.connect(self.__OpenFileDialog)
        self.Class_comboBox.currentTextChanged.connect(self.__get_current_class)
        self.Remove_Students_button.clicked.connect(self.__add_remove_list)
        self.Delete_Students_button.clicked.connect(self.__delete_student)
        self.Generate_button.clicked.connect(self.__SaveFileDialog)

    def __OpenFileDialog(self):
        if self.sender().objectName() == "SelectExcelPath_button":
            File_Path, _ = QFileDialog.getOpenFileName(self, filter = "Excel (*.xlsx)")
            self.ExcelPath_textEdit.setText(File_Path)
        elif self.sender().objectName() == "SelectPDFPath_button":
            File_Path, _ = QFileDialog.getOpenFileName(self, filter = "PDF (*.pdf)")
            self.PDFPath_textEdit.setText(File_Path)

        if self.ExcelPath_textEdit.toPlainText() and self.PDFPath_textEdit.toPlainText():
            ExcelPath = self.ExcelPath_textEdit.toPlainText()
            PDFPath = self.PDFPath_textEdit.toPlainText()
            self.ExcelFile = ReadWriteExcel(ExcelPath, PDFPath)

            self.Class_comboBox.clear()
            self.ExcelFile.TotalAmount = 0
            self.Remove_Students_comboBox.clear()

            if self.ExcelFile.SerialNumber and self.ExcelFile.Activityname and self.ExcelFile.Class_list:
                self.__add_info()
        else:
            self.SerialNumber_lineEdit.clear()
            self.Class_comboBox.clear()
            self.Activityname_lineEdit.clear()

    def __SaveFileDialog(self):
        if self.SerialNumber_lineEdit.text() and self.Activityname_lineEdit.text() and self.Class_comboBox.count() > 0:
            SaveDir = QFileDialog.getExistingDirectory()
            if SaveDir:
                self.ExcelFile.Generate(SaveDir, self.Class)
                self.Msg_textEdit.clear()
                for Message in self.ExcelFile.Message_log_list:
                    self.Msg(str(Message))
                self.ExcelFile.Message_log_list.clear()

    def __add_info(self):
        self.SerialNumber_lineEdit.setText(self.ExcelFile.SerialNumber)
        self.Activityname_lineEdit.setText(self.ExcelFile.Activityname)
        self.Class_comboBox.addItems(self.ExcelFile.Class_list)
        self.TotalAmount_label.setText(f"學生 {self.ExcelFile.TotalAmount} 位")
        self.__get_current_class()
        self.Remove_Students_comboBox.addItems(self.ExcelFile.Studentlist)

    def __clear_data(self):
        self.Remove_Students_comboBox.clear()
        self.Remove_Students_listWidget.clear()
        self.Remove_Students_list.clear()
        self.Msg_textEdit.clear()

    def __auto_add_remove_list(self): #自動新增學生至排除名單
        try:
            for index, Student in enumerate(self.ExcelFile.Studentlist):
                if int(self.ExcelFile.Hourslist[index]) == 0:
                    self.Remove_Students_listWidget.addItem(Student)
                    self.Remove_Students_list.append(Student)
                    self.Msg(f"已自動將 {Student} 新增至排除名單")
                    self.ExcelFile.Remove_Students_list = self.Remove_Students_list
        except ValueError:
            self.Msg(f"[錯誤] 研習時數資料型態有誤! 請確認Excel資料是否有誤")

    def __get_current_class(self):
        self.__clear_data()
        self.Class = str(self.Class_comboBox.currentText())
        self.TotalAmount, self.Studentlist = self.ExcelFile.get_student_list(self.Class)
        self.TotalAmount_label.setText(f"學生 {self.ExcelFile.TotalAmount} 位")
        self.Remove_Students_comboBox.addItems(self.Studentlist)
        self.__auto_add_remove_list()

    def __add_remove_list(self): #新增學生至排除名單
        CurrentStudent = self.Remove_Students_comboBox.currentText()
        if not CurrentStudent == "":
            if not CurrentStudent in self.Remove_Students_list:
                self.Remove_Students_listWidget.addItem(CurrentStudent)
                self.Remove_Students_list.append(CurrentStudent)
                self.Msg(f"已將 {CurrentStudent} 新增至排除名單")
            else:
                self.Msg(f"{CurrentStudent} 已在排除名單")

            self.ExcelFile.Remove_Students_list = self.Remove_Students_list
    
    def __delete_student(self): #將學生從排除名單移除
        if self.Remove_Students_list and self.Remove_Students_listWidget.selectedItems():
            DeleteStudent = self.Remove_Students_listWidget.currentItem().text()
            self.Remove_Students_list.remove(DeleteStudent)
            self.Remove_Students_listWidget.takeItem(self.Remove_Students_listWidget.currentIndex().row())
            self.Msg(f"已從排除名單移除 {DeleteStudent}")

            self.ExcelFile.Remove_Students_list = self.Remove_Students_list

    def Msg(self, Message: str):
        self.Msg_textEdit.append(f"{Message}")

class ReadWriteExcel():
    Sh1Title = []
    Sh1Values = []

    def __init__(self, ExcelPath: str, PDFPath: str):
        self.ExcelPath = ExcelPath
        self.PDFPath = PDFPath
        xlsx1 = pandas.read_excel(self.ExcelPath)
        self.Sh1Title = xlsx1.columns
        self.Sh1Values = xlsx1.values
        self.Sh1Titles = list(self.Sh1Title)
        self.Message_log_list = []
    
        self.SerialNumberIndex = self.Sh1Titles.index("招生編號")
        self.ActivitynameIndex = self.Sh1Titles.index("研習活動名稱")
        self.SchoolIndex = self.Sh1Titles.index("學校")
        self.ClassIndex = self.Sh1Titles.index("班級")
        self.StudentIndex = self.Sh1Titles.index("姓名")
        self.HoursIndex = self.Sh1Titles.index("研習時數")

        self.Remove_Students_list = []

        self.__get_info()
    
    def __get_info(self):
        self.SerialNumber = None
        self.Activityname = None

        self.SerialNumber = str(self.Sh1Values[0][self.SerialNumberIndex])
        self.Activityname = str(self.Sh1Values[0][self.ActivitynameIndex])
        self.Class_list = []
        for Sh1Value in self.Sh1Values:
            if not str(Sh1Value[self.ClassIndex]) in self.Class_list and not str(Sh1Value[self.ClassIndex]) == "nan":
                self.Class_list.append(str(Sh1Value[self.ClassIndex]))

    def get_student_list(self, Class): #獲取學生名單列表
        self.Studentlist = []
        self.StudentClasslist = []
        self.TotalAmount = 0
        self.Hourslist = []
        for Sh1Value in self.Sh1Values:
            if str(Sh1Value[self.SerialNumberIndex]) == str(self.SerialNumber) and \
                str(Sh1Value[self.ActivitynameIndex]) == str(self.Activityname) and \
                str(Sh1Value[self.ClassIndex]) == str(Class):
                    self.TotalAmount = self.TotalAmount + 1 #學生數量
                    self.Studentlist.append(Sh1Value[self.StudentIndex]) #學生列表
                    self.StudentClasslist.append(f"{Sh1Value[self.SchoolIndex]}{Sh1Value[self.ClassIndex]}") #學生班級
                    self.Hourslist.append(Sh1Value[self.HoursIndex]) #研習時數

        return self.TotalAmount, self.Studentlist

    def __split_pdf(self, SavePath): #分割pdf
        count = 0 #生成證書的數量
        remove_count = 0 #已排除的數量
        not_found_count = 0 #找不到學生的數量
        fileReader = PyPDF2.PdfReader(self.PDFPath) #讀取研習證書pdf
        
        for index in range(fileReader._get_num_pages()):
            page = fileReader.pages[index]
            text = page.extract_text()
            text_list = list(filter(None, text.split(" ")))
            name_index = next((i for i, item in enumerate(text_list) if "同學" in item), -1) #學生姓名的索引值
            student_name = text_list[name_index - 1] #當前頁面中的學生姓名
            student_class = text_list[name_index - 2][-2:] #獲取班級

            if str(student_name) in str(self.Studentlist) and student_class in self.StudentClasslist[0]: #判斷當前學生是否有在名單內
                if not str(student_name) in self.Remove_Students_list: #若學生不在移除名單內則生成pdf
                    fileWrite = PyPDF2.PdfWriter()
                    fileWrite.add_page(page)
                    Filename = f"{student_name}研習證書.pdf"
                    SaveFile_Path = f"{SavePath}/{Filename}"
                    fileWrite.write(SaveFile_Path)
                    count = count + 1
                    self.Message_log_list.append(f'{index + 1}. 已生成 {student_name}研習證書.pdf')
                else:
                    remove_count = remove_count + 1
                    self.Message_log_list.append(f"{index + 1}. 已排除 {student_name} 學生")
            else:
                not_found_count = not_found_count + 1
                self.Message_log_list.append(f'{index + 1}. {self.PDFPath} 中查無 {student_name}')
        self.Message_log_list.append(f"共 {fileReader._get_num_pages()} 頁，已生成 {count} 個證書，已排除 {remove_count} 位學生，查無 {not_found_count} 位學生")
        if count > 0:
            self.Message_log_list.append(f'已將檔案保存至 "{SavePath}"')
        else:
            if not os.listdir(SavePath):
                os.removedirs(SavePath)

    def Generate(self, SaveDir: str, Class: str): #生成檔案
        try:
            SavePath = f"{SaveDir}/{Class}" #存檔路徑
            
            if not os.path.isdir(SavePath):
                os.mkdir(f"{SavePath}")

            self.__split_pdf(SavePath)
            
        except Exception as ErrMsg:
            self.Message_log_list.append(f"錯誤: {ErrMsg}")
                

if __name__=='__main__':
    app = QApplication(sys.argv)
    window = ShowUI()
    window.show()
    sys.exit(app.exec_())