import sys
from PyQt5 import QtWidgets
from openpyxl import load_workbook
from ui_generatingCertificate_v2 import Ui_MainWindow

class CertificateApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.TypeCertificate.clear()
        self.TypeCertificate.addItems(["Справка1", "Справка2", "Справка3", "Справка4"])
        self.PreviewButton.clicked.connect(self.printCertificate)

    def printCertificate(self):
        # Загружаем книгу Excel и активный лист
        workbook = load_workbook('database.xlsx')
        sheet = workbook.active

        # Словарь заголовков и соответствующих значений для записи
        data = {
            'TypeCertificate': self.TypeCertificate.currentText(),
            'NumberTicket': self.NumberTicket.text(),
            'Itinerary': self.Itinerary.text(),
            'LastName': self.LastNameLineEdit.text(),
            'Name': self.NameLineEdit.text(),
            'Surname': self.SurnameLineEdit.text(),
            'Passport': self.PassportLineEdit.text(),
            'IssuingAuthority': self.PassportInformation.text(),
            'Telephone': self.Telephone.text(),
            'Email': self.Email.text(),
            'RecipientName': self.nameRecipient.text(),
        }

        # Сопоставляем колонки Excel и данные с формы
        headers = [cell.value for cell in sheet[1]]  # Предполагаем, что заголовки находятся в первой строке
        next_row = sheet.max_row + 1
        for header in headers:
            col = headers.index(header) + 1  # Получаем индекс заголовка для определения номера столбца
            sheet.cell(row=next_row, column=col, value=data[header])

        # Сохраняем изменения
        workbook.save('database.xlsx')
        QtWidgets.QMessageBox.information(self, 'Успех', 'Данные добавлены в файл database.xlsx')

# Инициализация и запуск приложения
app = QtWidgets.QApplication(sys.argv)
mainWindow = CertificateApp()
mainWindow.show()
sys.exit(app.exec_())
