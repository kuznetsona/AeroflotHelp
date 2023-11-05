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
        # Открываем существующий файл Excel
        workbook = load_workbook('database.xlsx')
        sheet = workbook.active

        # Определяем номер следующей строки для записи данных
        next_row = sheet.max_row + 1

        # Считываем данные с формы
        data = {
            'Вид справки': self.TypeCertificate.currentText(),
            'Номер билета': self.NumberTicket.text(),
            'Маршрут следования': self.Itinerary.text(),
            'Фамилия': self.LastNameLineEdit.text(),
            'Имя': self.NameLineEdit.text(),
            'Отчество': self.SurnameLineEdit.text(),
            'Паспорт': self.PassportLineEdit.text(),
            'Выдающий орган': self.PassportInformation.text(),
            'Телефон': self.Telephone.text(),
            'Почта': self.Email.text(),
            'ФИО получателя': self.nameRecipient.text(),
        }

        # Записываем данные в новую строку
        for column, value in enumerate(data.values(), start=1):
            sheet.cell(row=next_row, column=column).value = value

        # Сохраняем изменения в файл Excel
        workbook.save('database.xlsx')
        QtWidgets.QMessageBox.information(self, 'Успех', 'Данные добавлены в файл database.xlsx')

# Инициализация и запуск приложения
app = QtWidgets.QApplication(sys.argv)
mainWindow = CertificateApp()
mainWindow.show()
sys.exit(app.exec_())
