import sys
import os
import wmi
import pythoncom
import ctypes  
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.uic import loadUi
from PyQt5.QtCore import QThread, pyqtSignal

class USBThread(QThread):
    usb_connected = pyqtSignal()

    def run(self):
        pythoncom.CoInitialize()  # Инициализация COM-объектов
        c = wmi.WMI()
        print("Отслеживание подключения USB устройств. Для выхода нажмите Ctrl+C.")
        while True:
            query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_USBHub'"
            watcher = c.watch_for(raw_wql=query)
            usb_event = watcher()
            print(f"USB устройство подключено: {usb_event}")
            
            print(c.Win32_LogicalDisk())
            self.usb_connected.emit()

class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        loadUi('main.ui', self)
        self.show()
        
        self.pushButton.clicked.connect(self.update_device_info) 
        self.pushButton_2.clicked.connect(self.save_text_to_file_extensions)
        self.pushButton_3.clicked.connect(self.save_text_to_file_serial_numbers)
        self.update_device_info()

        self.usb_thread = USBThread()
        self.usb_thread.usb_connected.connect(self.update_device_info)
        self.usb_thread.start()

        # Скрыть вкладку для обычных пользователей
        if not self.is_admin():
            self.tabWidget.setTabEnabled(1, False)
        else:
            extensions_text = self.read_text_from_file("extensions.txt")
            self.plainTextEdit_1.setPlainText(extensions_text)
            serial_numbers_text = self.read_text_from_file("serial_numbers.txt")
            self.plainTextEdit_2.setPlainText(serial_numbers_text)

    def save_text_to_file_extensions(self):
        text_to_save = self.plainTextEdit_1.toPlainText()
        with open("extensions.txt", "w") as file:
            file.write(text_to_save)
        QMessageBox.information(self, "Уведомление", "Расширения успешно сохранены в файл extensions.txt")

    def save_text_to_file_serial_numbers(self):
        text_to_save = self.plainTextEdit_2.toPlainText()
        with open("serial_numbers.txt", "w") as file:
            file.write(text_to_save)
        QMessageBox.information(self, "Уведомление", "Список СМНИ успешно сохранен в файл serial_numbers.txt")

    def read_text_from_file(self, filename):
        try:
            with open(filename, "r") as file:
                text = file.read()
            return text
        except FileNotFoundError:
            return "Файл не найден"

    def is_admin(self):
        # Проверка на администратора
        try:
            return os.getuid() == 0
        except AttributeError:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0

    def update_device_info(self):
        c = wmi.WMI()
        usb_info = ""
        forbidden_extensions = self.load_extensions()
        forbidden_serials = self.load_serial_numbers()
        i = 0
        for drive in c.Win32_LogicalDisk():
            usb_info += f"Метка устройства: {drive.DeviceID}<br>"
            
            for info_drive in c.Win32_DiskDrive():
                if i == info_drive.Index:
                    usb_info += f"Описание устройства: {info_drive.Description}<br>"
                    usb_info += f"Тип устройства: {info_drive.InterfaceType}<br>"
                    usb_info += f"Модель устройства: {info_drive.Model}<br>"
                    usb_info += f"Серийный номер: {info_drive.SerialNumber}<br>"
                    usb_info_last = ""
                    if info_drive.InterfaceType == "USB":
                        usb_info_last += f"Метка устройства: {drive.DeviceID}<br>"
                        usb_info_last += f"Серийный номер: {info_drive.SerialNumber}<br>"
                        usb_files = self.get_usb_files(drive.DeviceID)
                        name_flesh = drive.DeviceID
                        serial_number = info_drive.SerialNumber
                        if self.check_serial_number(forbidden_serials, serial_number):
                            if self.check_extension(forbidden_extensions, usb_files):
                                usb_info += "<font color='red'>На устройстве обнаружены запрещённые файлы</font><br>"
                                usb_info_last += "<font color='red'>На устройстве обнаружены запрещённые файлы</font><br>"
                                usb_info_last += "<font color='red'>Устройство заблокировано!</font><br>"
                                os.system(f'powershell $driveEject = New-Object -comObject Shell.Application; $driveEject.Namespace(17).ParseName("""{name_flesh}""").InvokeVerb("""Eject""")')
                            else:
                                usb_info += "<font color='green'>Устройство в норме</font><br>"
                                usb_info_last += "<font color='green'>Устройство прошло проверку</font><br>" 
                                usb_info_last += "<font color='green'>Доступ открыт!</font><br>"
                        else:    
                            usb_info += "<font color='red'>Данное устройство не найдено в перечне СМНИ</font><br>"
                            usb_info_last += "<font color='red'>Устройство не найдено в перечне СМНИ</font><br>"
                            usb_info_last += "<font color='red'>Устройство заблокировано!</font><br>"
                            os.system(f'powershell $driveEject = New-Object -comObject Shell.Application; $driveEject.Namespace(17).ParseName("""{name_flesh}""").InvokeVerb("""Eject""")')

                    else:
                        usb_info += "<font color='green'>Устройство в норме</font><br>"
            usb_info += "<br><br>"
            i += 1
        self.textBrowser_2.setHtml(usb_info)    
        self.textBrowser_2.verticalScrollBar().setValue(0)
        self.textBrowser.setHtml(usb_info_last)

    def get_usb_files(self, device_id):
        usb_files = []
        for root, dirs, files in os.walk(device_id):
            for file in files:
                usb_files.append(os.path.join(root, file))
        return usb_files

    def load_extensions(self):
        with open("extensions.txt", "r") as file:
            extensions = [line.strip() for line in file]
        return extensions

    def load_serial_numbers(self):
        with open("serial_numbers.txt", "r") as file:
            serial_numbers = [line.strip() for line in file]
        return serial_numbers

    def check_extension(self, forbidden_extensions, usb_files):
        for file in usb_files:
            extension = os.path.splitext(file)[1]
            if extension in forbidden_extensions:
                return True
        return False

    def check_serial_number(self, forbidden_serials, serial_number):
        return serial_number in forbidden_serials


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    sys.exit(app.exec_())
