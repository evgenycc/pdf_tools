"""
Небольшое приложение с GUI, которое выполняет конвертацию
локальных страниц html в pdf, объединяет pdf в один файл, а также
пытается уменьшить размер указанного файла pdf.
Требует для установки следующие библиотеки:
pip install pyqt5 PyPDF2 pdfkit PyMuPDF,
а также необходимо установить пакет для конвертации html:
sudo apt-get install wkhtmltopdf

Данный код предназначен в первую очередь для ОС linux, так как
для Windows приложений подобного рода хватает. Если возникнет необходимость
использовать данный код в ОС Windows, нужно будет изменить модуль для вывода
всплывающих сообщений, установить GhostScript, а также wkhtmltopdf.
С последним эксперименты не проводил, потому не знаю, как он будет работать
в Windows. Однако, есть примеры, в которых можно указать путь к файлу exe и
все будет работать.

В Linux также есть возможность указать путь к исполняемому файлу, но лучше
установить пакет в систему. Так как, в случае использования не установленного
пакета он работает, но вероятность ошибок при конвертации возрастает.
"""

import os.path
import shutil
import subprocess
import sys
from os import system
from pathlib import Path

import fitz
import pdfkit
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyQt5.QtWidgets import QLineEdit

from mergepdf import *


class MyWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.path_file = None
        self.ui.mergeBtn.setEnabled(False)
        self.ui.convertBtn.setEnabled(False)
        self.ui.compressBtn.setEnabled(False)
        self.ui.ratioBox.setCurrentText('3')
        self.ui.fileView.horizontalHeader().resizeSection(0, 895)
        self.ui.folderBtn.clicked.connect(self.folder_open)
        self.ui.mergeBtn.clicked.connect(self.merge_pdf)
        self.ui.convertBtn.clicked.connect(self.convert_html)
        self.ui.delBtn.clicked.connect(self.del_line)
        self.ui.upBtn.clicked.connect(self.move_up)
        self.ui.dwnBtn.clicked.connect(self.move_down)
        self.ui.compressBtn.clicked.connect(self.compress_pdf)
        self.ui.methodBox.currentTextChanged.connect(self.change_method)
        self.ui.exitBtn.clicked.connect(self.exit_application)

    def exit_application(self):
        """
        Завершение работы приложения.
        """
        QtWidgets.QApplication.quit()

    def move_down(self):
        """
        Перемещение строки в таблице вниз.
        Определяем индекс текущей строки и столбца. Сравниваем текущий
        индекс строки. Если он меньше, чем текущий индекс - 1, создаем
        новую строку. Запускаем цикл в диапазоне количества столбцов, перемещаем
        содержимое вниз. После цикла удаляем строку, которую нужно было переместить.
        """
        row = self.ui.fileView.currentRow()
        column = self.ui.fileView.currentColumn()
        if row < self.ui.fileView.rowCount() - 1:
            self.ui.fileView.insertRow(row + 2)
            for i in range(self.ui.fileView.columnCount()):
                self.ui.fileView.setItem(row + 2, i, self.ui.fileView.takeItem(row, i))
                self.ui.fileView.setCurrentCell(row + 2, column)
            self.ui.fileView.removeRow(row)

    def move_up(self):
        """
        Перемещение строки в таблице вверх. Повторяем все те же операции, но сравниваем
        индекс текущей строки с нулем. Пока он больше, можем перемещать.
        """
        row = self.ui.fileView.currentRow()
        column = self.ui.fileView.currentColumn()
        if row > 0:
            self.ui.fileView.insertRow(row - 1)
            for i in range(self.ui.fileView.columnCount()):
                self.ui.fileView.setItem(row - 1, i, self.ui.fileView.takeItem(row + 1, i))
                self.ui.fileView.setCurrentCell(row - 1, column)
            self.ui.fileView.removeRow(row + 1)

    def del_line(self):
        """
        Удаление строки из таблицы.
        Определяем индекс строки, удаляем строку из таблицы
        по индексу.
        """
        row = self.ui.fileView.currentRow()
        self.ui.fileView.removeRow(row)

    def convert_html(self):
        """
        Функция для конвертации локальных страниц html
        в pdf. Устанавливаем значение прогресс-бара на 0.
        Создаем список из значений строк таблицы, то выбираем
        только те записи, которые имеют расширение html.
        Проверяем длину списка. Если 0, выходим из функции,
        в статус-бар пишем сообщение о том, что конвертировать нечего,
        то же сообщение выводим на экран.
        Если длина списка больше 0, устанавливаем максимальное значение
        статус-бара на длину списка, запускаем цикл по списку, с помощью функции
        enumerate получаем индекс элемента, устанавливаем значение статус-бара равным
        индексу + 1.
        Формируем пути к входным html и выходным pdf файлам.
        Запускаем конвертацию из файла, в качестве параметров передаем путь к файлу
        html и путь для выходного файла pdf. Обрабатываем исключение.
        После завершения цикла выводим в статус-бар сообщение о полной конвертации,
        то же сообщение выводим на экран. Запускаем функцию для чтения файлов из
        директории и переформирования содержимого таблицы на экране.
        :return: Выход из функции в случае ошибки.
        """
        self.ui.progressBar.setValue(0)
        files_html = [self.ui.fileView.item(path, 0).text() for path in range(0, self.ui.fileView.rowCount())
                      if self.ui.fileView.item(path, 0).text().endswith(".html")]

        if len(files_html) == 0:
            self.ui.statusbar.showMessage('Нечего конвертировать')
            linux_notify(f'"Нечего конвертировать"')
            return

        self.ui.progressBar.setMaximum(len(files_html))
        for num, file in enumerate(files_html):
            self.ui.progressBar.setValue(num + 1)
            html_name = os.path.join(self.path_file, file)
            pdf_name = os.path.join(self.path_file, f'{file[0:-4]}pdf')
            try:
                pdfkit.from_file(html_name, pdf_name)
            except OSError as ex:
                print(ex)
                continue
        self.ui.statusbar.showMessage('Все файлы конвертированы')
        linux_notify(f'"Все файлы конвертированы"')
        self.file_read()

    def merge_pdf(self):
        """
        Функция для слияния файлов pdf из таблицы. Слияние выполняется
        с помощью fitz. Создаем объект fitz, открываем каждый из файлов
        pdf и добавляем в созданные объект. После чего выполняем сохранение.
        Работает в разы быстрее аналогичной функции слияния у PyPDF2.
        :return: Выход из функции.
        """
        self.ui.progressBar.setValue(0)
        files_pdf = [self.ui.fileView.item(path, 0).text() for path in range(0, self.ui.fileView.rowCount())
                     if self.ui.fileView.item(path, 0).text().endswith("pdf")]
        if len(files_pdf) == 0 or len(files_pdf) == 1:
            self.ui.statusbar.showMessage('Нечего объединять')
            linux_notify(f'"Нечего объединять"')
            return
        result = fitz.open()
        path_dir = self.ui.pathEdt.text()
        save_path = QtWidgets.QFileDialog.getSaveFileName(None, "Имя для объединенного файла", path_dir, "*.pdf")
        path_name = f'{save_path[0]}{save_path[1].replace("*", "")}'
        if path_name == '':
            return

        self.ui.progressBar.setMaximum(len(files_pdf))
        for num, pdf in enumerate(files_pdf):
            file = os.path.join(self.path_file, pdf)
            self.ui.progressBar.setValue(num)
            with fitz.open(file) as mfile:
                result.insert_pdf(mfile)
        result.save(path_name)
        self.ui.progressBar.setValue(len(files_pdf))
        self.ui.statusbar.showMessage('Объединение завершено')
        linux_notify(f'"Объединение завершено"')
        self.file_read()

    def change_method(self):
        """
        Проверка метода сжатия на форме. В зависимости от этого
        делаем комбо-бокс со степенью сжатия активным или нет.
        :return: ВЫход из функции.
        """
        if self.ui.methodBox.currentText() == 'Ghost Script':
            self.ui.ratioBox.setEnabled(True)
        else:
            self.ui.ratioBox.setEnabled(False)

    def compress_pdf(self):
        """
        Обработка значения в комбо-боксе с методом сжатия.
        В зависимости от выбранного метода запускается необходимая
        для выполнения этого метода функци.
        :return: Выход из функции.
        """
        row = self.ui.fileView.currentRow()
        try:
            if not self.ui.fileView.item(row, 0).text().endswith(".pdf"):
                self.ui.statusbar.showMessage('Выбранный файл не PDF')
                linux_notify(f'"Выбранный файл не PDF"')
                return
        except AttributeError:
            self.ui.statusbar.showMessage('Файл не выбран')
            linux_notify(f'"Файл не выбран"')
            return

        path_file = os.path.join(self.path_file, self.ui.fileView.item(row, 0).text())
        out_file = f'{path_file[0:-4]}_compress.pdf'

        if self.ui.methodBox.currentText() == 'Ghost Script':
            self.ui.statusbar.showMessage('Сжатие PDF')
            linux_notify(f'"Сжатие PDF"')
            compress_perc = self.compress_gs(path_file, out_file)
            self.ui.statusbar.showMessage(compress_perc)
            linux_notify(f'"{compress_perc}"')
            self.file_read()
        elif self.ui.methodBox.currentText() == 'PyPDF2':
            self.ui.statusbar.showMessage('Сжатие PDF')
            linux_notify(f'"Сжатие PDF"')
            compress_perc = self.compress_file_pypdf2(path_file, out_file)
            self.ui.statusbar.showMessage(compress_perc)
            linux_notify(f'"{compress_perc}"')
            self.file_read()

    def compress_gs(self, path_file, out_file):
        """
        Сжатие файла pdf с помощью GhostScript.
        Параметр качества берется непосредственно из формы.
        Проверяется, существует ли файл, получаем путь к
        GhostScript. Считываем размер файла до сжатия.
        Запускаем команду для сжатия файла с передачей необходимых параметров.
        Считываем размер файла после сжатия, считаем коэффициент.
        :param path_file: путь к входному файлу.
        :param out_file: путь для выходного файла.
        :return: Возвращает степень сжатия файла в процентах.
        """
        power = int(self.ui.ratioBox.currentText())
        quality = {
            0: '/default',
            1: '/prepress',
            2: '/printer',
            3: '/ebook',
            4: '/screen'
        }

        if not os.path.isfile(path_file):
            return "Ошибка: неверный путь к файлу PDF"

        gs = shutil.which('gs')
        initial_size = os.path.getsize(path_file)
        subprocess.call([gs, '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                         f'-dPDFSETTINGS={quality[power]}',
                         '-dNOPAUSE', '-dQUIET', '-dBATCH',
                         f'-sOutputFile={out_file}',
                         path_file])
        final_size = os.path.getsize(out_file)
        ratio = 1 - (final_size / initial_size)
        return f"Compression by {ratio:.0%}"

    def compress_file_pypdf2(self, path_file, out_file):
        """
        Функция сжатия файла pdf с помощью библиотеки
        PyPDF2.
        Получаем размер файла до сжатия, создаем экземпляр
        класса писателя. Создаем экземпляр класса чтеца, с передачей
        пути к файлу для сжатия.
        Устанавливаем максимальное значение прогресс-бара.
        Запускаем цикл в диапазоне равном кол-ву страниц.
        Считываем каждую страницу, сжимаем, передаем писателю.
        Записываем массив данных страниц в файл.
        :param path_file: путь к входному файлу.
        :param out_file: путь для выходного файла.
        :return: Возвращает степень сжатия в процентах.
        """
        self.ui.progressBar.setValue(0)
        start_size = os.path.getsize(path_file)
        writer = PdfFileWriter()
        reader = PdfFileReader(path_file)
        self.ui.progressBar.setMaximum(reader.numPages)
        for num in range(0, reader.numPages):
            self.ui.progressBar.setValue(num+1)
            page = reader.getPage(num)
            page.compressContentStreams()
            writer.addPage(page)
        with open(out_file, 'wb') as file:
            writer.write(file)

        compress_size = os.path.getsize(out_file)
        compression_ratio = 1 - (compress_size / start_size)
        return f'Compression by: {compression_ratio:.0%}'

    def file_read(self):
        """
        Считываем содержимое полученной ранее директории с файлами.
        Добавляем в список все файлы pdf и html.
        Запускаем цикл по отсортированному списку, добавляем названия файлов
        таблицу, очищаем список файлов.
        """
        files = [x for x in Path(self.path_file).iterdir() if x.suffix == ".pdf" or x.suffix == ".html"]
        self.ui.fileView.setRowCount(0)

        if len(files) == 0:
            self.ui.statusbar.showMessage('Нет файлов для обработки')
            linux_notify('"Нет файлов для обработки"')
            return

        for fil in sorted(files):
            row_position = self.ui.fileView.rowCount()
            self.ui.fileView.insertRow(row_position)
            self.ui.fileView.setItem(row_position, 0, QtWidgets.QTableWidgetItem(str(fil.name)))
        files.clear()

    def folder_open(self):
        """
        Функция для получения пути к директории для открытия.
        Устанавливаем значения статус-бара, прогресс-бара.
        Получаем из диалога путь к директории, добавлем его в
        текстовую строку. Проверяем, не является ли полученное значение
        пустым, так как пользователь может передумать и нажать кнопку отмена.
        Если нет, активируем кнопки конвертации, слияния и сжатия.
        Запускаем функцию чтения содержимого директории.
        :return: Выход из функции.
        """
        self.ui.statusbar.showMessage('')
        self.ui.progressBar.setMaximum(100)
        self.ui.progressBar.setValue(0)

        self.path_file = QtWidgets.QFileDialog.getExistingDirectory(None, "Выбрать папку", ".")
        self.ui.pathEdt.setText(self.path_file)

        if self.path_file != '':
            self.file_read()
            self.ui.mergeBtn.setEnabled(True)
            self.ui.convertBtn.setEnabled(True)
            self.ui.compressBtn.setEnabled(True)
            self.ui.methodBox.setEnabled(True)
        else:
            self.ui.statusbar.showMessage('Не выбрана директория с файлами')
            linux_notify(f'"Не выбрана директория с файлами"')
            self.ui.pathEdt.setText('')
            self.ui.folderBtn.setFocus()
            return


def linux_notify(text):
    """
    Функция выводит сообщение с текстом на экран.
    Получаем текст, формируем команду, выполняем.
    :param text: Текстовая строка для сообщения.
    """
    command = f'''notify-send {text}'''
    system(command)


def update():
    """
    Проверка наличия необходимых пакетов, для работы скрипта.
    Проверяем, есть ли в переменных окружения пакет wkhtmltopdf.
    Необходим для конвертации html в pdf.
    Если данного пакета нет, выводим сообщение пользователю с просьбой
    ввести пароль sudo для установки данного пакета.
    Передаем параметры в команду Popen, а также введенных пароль.
    То же самое для GhostScript. Если в системе нет данного модуля, пользователю
    будет предложено его установить.
    """
    if not shutil.which('wkhtmltopdf'):
        sudo_password = QtWidgets.QInputDialog.getText(None, "Пароль для установки", "Внимание!\nНе установлен"
                                                                                     " wkhtmltopdf.\n"
                                                                                     "Для его установки введите "
                                                                                     "пароль sudo:",
                                                       QLineEdit.Normal)[0]
        command1 = 'apt install wkhtmltopdf'.split()
        p = subprocess.Popen(['sudo', '-S'] + command1, stdin=subprocess.PIPE, stderr=subprocess.PIPE,
                             universal_newlines=True)
        p.communicate(sudo_password + '\n')

    if not shutil.which('gs'):
        sudo_password = QtWidgets.QInputDialog.getText(None, "Пароль для установки", "Внимание!\nНе установлен"
                                                                                     " Ghost Script.\n"
                                                                                     "Для его установки введите "
                                                                                     "пароль sudo:",
                                                       QLineEdit.Normal)[0]
        command1 = 'apt install ghostscript'.split()
        p = subprocess.Popen(['sudo', '-S'] + command1, stdin=subprocess.PIPE, stderr=subprocess.PIPE,
                             universal_newlines=True)
        p.communicate(sudo_password + '\n')


if __name__ == '__main__':
    """
    Создается объект с Qt виджетами, запускается функция проверки пакетов,
    запускается Qt приложение.
    """
    app = QtWidgets.QApplication(sys.argv)
    update()
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())
