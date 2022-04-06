from PyQt5 import QtWidgets, QtGui, QtCore
import design_editing_text  # конвертированный файл дизайна
import re
import sys, os, inspect
import psutil, subprocess
import win32gui, win32con
import webbrowser
import time
sys.path.append("..")
from audio_helper import MyRecogn

class AudioRecognitionThread(QtCore.QThread):
    def __init__(self, parent=None):
        super(AudioRecognitionThread, self).__init__(parent)
        self.main_text = ''
        self.my_rec = MyRecogn()
        self.mes_init = self.my_rec.mes_init

    def run(self):
        while True:
            query = self.my_rec.listen_command(self.signal_status)
            self.main_text = query
            self.my_signal.emit(self.main_text)

class EditTextApp(QtWidgets.QMainWindow, design_editing_text.Ui_Dialog):
    my_signal = QtCore.pyqtSignal(str, name='my_signal')
    signal_status = QtCore.pyqtSignal(str, name='signal_status')
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.commands_dict = {
            'commands': {
                'show_window': ['Масяня', 'на сегодня'],
                'hide_window': ['скройся', 'закройся', 'свернись'],
                'start_messengers': ['режим сообщения', 'режим сообщений', 'режим сообщение'],
                'stop_recording': ['Остановить запись', 'остановить запись']
            }
        }
        self.messengers = ["WhatsApp.exe", "Telegram.exe"]
        self.time_mode_silence = 0  # таймер для отключения режима тишины (фриз мессенджеров)


        self.pushButtonStopRec.clicked.connect(self.pressButtonStopRec)
        self.pushButton.clicked.connect(self.pressButtonCopyToClipboard)
        self.pushButtonBack.clicked.connect(self.pressButtonBack)
        self.pushButtonForward.clicked.connect(self.pressButtonForward)
        self.pushButtonHelp.clicked.connect(self.pressButtonHelp)
        self.my_signal.connect(self.mySignalHandler)
        self.signal_status.connect(self.SignalStatusRecognition)
        # for test
        # self.textBrowser.insertPlainText('Test.')
        self.pressButtonStopRec()

    @QtCore.pyqtSlot(str)
    def SignalStatusRecognition(self, status):
        self.label_status.setText(status)

    @QtCore.pyqtSlot(str)
    def mySignalHandler(self, text_recogn):
        if text_recogn:
            self.textBrowserLastRecogn.setText(text_recogn)
            fi1 = text_recogn.find(' заменить на ')
            if fi1 > -1:
                patt = text_recogn[0: fi1]
                cur_text = self.textBrowser.toPlainText()
                cur_text = cur_text.replace(patt, text_recogn[fi1 + 13:])
                self.textBrowser.setText(cur_text)

            elif text_recogn == 'маленькими буквами' or text_recogn == 'маленькие буквы':
                cursor = self.textBrowser.textCursor()
                self.textBrowser.insertPlainText(cursor.selectedText().lower())

            elif text_recogn == 'большими буквами' or text_recogn == 'большие буквы':
                cursor = self.textBrowser.textCursor()
                self.textBrowser.insertPlainText(cursor.selectedText().upper())

            elif text_recogn == 'с большой буквы' or text_recogn == 'большая буква':
                cursor = self.textBrowser.textCursor()
                sel_text = cursor.selectedText()
                sel_text = sel_text[0].upper() + sel_text[1:]
                self.textBrowser.insertPlainText(sel_text)

            elif text_recogn == 'скопировать в буфер' or text_recogn == 'копировать в буфер':
                self.pressButtonCopyToClipboard()

            elif text_recogn == 'загуглить':
                self.google_it(self.textBrowser.toPlainText())

            elif text_recogn == 'очистить всё' or text_recogn == 'Удалить всё':
                self.textBrowser.setText('')

            elif text_recogn == 'удалить':
                cursor = self.textBrowser.textCursor()
                if cursor.selectionStart():
                    # удаляем пробельный символ перед выделением, чтобы избегать накопления пробелов после удаления
                    if self.textBrowser.toPlainText()[cursor.selectionStart() - 1] == ' ':
                        end_symb = cursor.selectionEnd()
                        cursor.setPosition(cursor.selectionStart() - 1)
                        cursor.setPosition(end_symb, QtGui.QTextCursor.KeepAnchor)
                        self.textBrowser.setTextCursor(cursor)
                        cursor = self.textBrowser.textCursor()
                cursor.removeSelectedText()

            elif text_recogn in self.commands_dict['commands']['show_window']:
                self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                self.show()
                self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
                self.show()

            elif text_recogn in self.commands_dict['commands']['hide_window']:
                self.hide()

            elif text_recogn in self.commands_dict['commands']['stop_recording']:
                self.pressButtonStopRec()

            elif text_recogn == 'режим тишины':
                self.turn_off_messengers()
            elif re.search("режим тишины на (\d+) минут", text_recogn):
                minutes = int(re.search("режим тишины на (\d+) минут", text_recogn).group(1))
                self.time_mode_silence = time.time() + 60 * minutes
                print('Мессенджеры временно заморожены. Их работа будет возобновлена ' +
                      time.asctime(time.localtime(self.time_mode_silence)))
                self.turn_off_messengers()

            elif text_recogn in self.commands_dict['commands']['start_messengers']:
                self.turn_on_messengers()

            elif text_recogn == 'фокус WhatsApp':
                self.show_window('WhatsApp')
            elif text_recogn == 'фокус телеграм':
                self.show_window('Telegram (')

            elif text_recogn == 'действие назад':
                self.pressButtonBack()
            elif text_recogn == 'действие вперед':
                self.pressButtonForward()

            elif text_recogn == 'воспроизвести запись':
                subprocess.Popen(r'c:\AutoClickExtreme\AutoClicker.exe -play c:\AutoClickExtreme\aips\whatsapp_focus.aip')

            else:
                # режим побуквенного ввода (аналог аббревиатур)
                abbr_mode = False
                if text_recogn[0: 21] == 'режим маленькие буквы':
                    abbr_mode = True
                    text_recogn = self.extract_letters(text_recogn[22:]).lower()
                elif text_recogn[0: 24] == 'режим маленькими буквами':
                    abbr_mode = True
                    text_recogn = self.extract_letters(text_recogn[25:]).lower()
                elif text_recogn[0: 19] == 'режим большие буквы':
                    abbr_mode = True
                    text_recogn = self.extract_letters(text_recogn[20:]).upper()
                elif text_recogn[0: 22] == 'режим большими буквами':
                    abbr_mode = True
                    text_recogn = self.extract_letters(text_recogn[23:]).upper()
                text_recogn = text_recogn.replace('точка', '.').replace('Точка', '.').replace(' .', '.')
                text_recogn = text_recogn.replace('запятая', ',').replace('Запятая', ',').replace(' ,', ',')
                text_recogn = text_recogn.replace('двоеточие', ':').replace('Двоеточие', ':').replace(' :', ':')
                text_recogn = text_recogn.replace('знак вопроса', '?').replace(' ?', '?')
                text_recogn = text_recogn.replace('восклицательный знак', '!').replace(' !', '!')
                text_recogn = text_recogn.replace('тире', ' - ').replace('  -', ' -')
                cursor = self.textBrowser.textCursor()
                cursor.beginEditBlock()
                if not cursor.position() and not abbr_mode:
                    # с заглавной буквы пишем начало текста
                    text_recogn = text_recogn[0].upper() + text_recogn[1:]
                else:
                    cur_txt = self.textBrowser.toPlainText()
                    # если точка перед вставляемым текстом, то делаем вставляемый текст с большой буквы
                    if not abbr_mode and cur_txt[cursor.selectionStart() - 1] == '.':
                        text_recogn = text_recogn[0].upper() + text_recogn[1:]
                    # добавляем пробел, если курсор стоит не после пробела (иначе диктуемое слово будет
                    # без пробелов присоединяться к предыдущему слову)
                    if not abbr_mode and re.search("\w", text_recogn[0]) and \
                            cur_txt[cursor.selectionStart() - 1] != ' ':
                        text_recogn = ' ' + text_recogn
                # после точки преобразуем маленькую букву в большую
                m = re.search("(\. [a-z,а-я])", text_recogn)
                if m:
                    text_recogn = text_recogn.replace(m.group(1), m.group(1).upper())
                self.textBrowser.insertPlainText(text_recogn)
                cursor.endEditBlock()
        if self.time_mode_silence and self.time_mode_silence < time.time():
            self.turn_on_messengers()

    def google_it(self, text):
        webbrowser.open_new('https://www.google.ru/search?q=' + text)

    def extract_letters(self, text_recogn):
        if not text_recogn:
            return ''
        text_recogn = text_recogn.replace('мягкий знак', 'ь').replace('твердый знак', 'ъ')
        res_abbr = text_recogn[0]
        fi = text_recogn.find(' ')
        while fi > -1:
            res_abbr += text_recogn[fi + 1]
            fi = text_recogn.find(' ', fi + 1)
        return res_abbr

    def pressButtonStopRec(self):
        if self.pushButtonStopRec.text() == 'Начать запись':
            self.audio_inst = AudioRecognitionThread()
            self.audio_inst.my_signal = self.my_signal
            self.audio_inst.signal_status = self.signal_status
            self.textBrowserLastRecogn.setText(self.audio_inst.mes_init)
            self.audio_inst.start()
            self.pushButtonStopRec.setText('Остановить запись')
        else:
            self.audio_inst.terminate()
            self.label_status.setText('Запись остановлена')
            self.pushButtonStopRec.setText('Начать запись')


    def pressButtonCopyToClipboard(self):
        clipboard = QtGui.QGuiApplication.clipboard()
        # originalText = clipboard.text()
        clipboard.setText(self.textBrowser.toPlainText())

    def pressButtonBack(self):
        self.textBrowser.undo()
    def pressButtonForward(self):
        self.textBrowser.redo()

    def pressButtonHelp(self):
        script_path = os.path.realpath(os.path.abspath(os.path.join(os.path.split(
            inspect.getfile(inspect.currentframe()))[0])))
        script_path = script_path.replace("\\", "/")
        subprocess.Popen('notepad %s/readme.txt' % script_path)

    def turn_on_messengers(self):
        for process in (process for process in psutil.process_iter()
                        if process.name() in self.messengers):
            process.resume()
        self.time_mode_silence = 0
        # subprocess.Popen(r'C:\Users\%s\AppData\Roaming\Telegram Desktop\Telegram.exe' % os.getlogin())
        # subprocess.Popen(r'C:\Users\%s\AppData\Local\WhatsApp\app-2.2202.12\WhatsApp.exe' % os.getlogin())

    def turn_off_messengers(self):
        for process in (process for process in psutil.process_iter()
                        if process.name() in self.messengers):
            process.suspend()
            # process.kill()

    def windowEnumerationHandler(self, hwnd, top_windows):
        top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
    def show_window(self, capt):
        top_windows = []
        win32gui.EnumWindows(self.windowEnumerationHandler, top_windows)
        count = 0
        for i in top_windows:
            if capt in i[1]:
                count += 1
                # win32gui.ShowWindow(i[0], win32con.SW_HIDE)
                # l, t, r, b = win32gui.GetWindowRect(i[0])
                # if self.second_monitor:
                #     win32gui.MoveWindow(i[0], 2000, 200, r - l, b - t, False)
                # else:
                #     win32gui.MoveWindow(i[0], 1000, 300, r - l, b - t, False)
                # win32gui.ShowWindow(i[0], win32con.SW_MINIMIZE)
                # win32gui.ShowWindow(i[0], win32con.SW_SHOWNOACTIVATE)
                # win32gui.ShowWindow(i[0], 5)
                try:
                    win32gui.ShowWindow(i[0], win32con.SW_MINIMIZE)
                    win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
                    win32gui.SetForegroundWindow(i[0])
                except:
                    pass
                # if count == 2:
                #     break

app = QtWidgets.QApplication([])
window = EditTextApp()
window.show()
sys.exit(app.exec_())
