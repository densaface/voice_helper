# Голосовой ассистент КЕША 1.0 BETA
import speech_recognition
# from fuzzywuzzy import fuzz
import pyttsx3
import datetime
import time, os
import pywintypes
import win32pipe, win32file

debug_level = 5

# настройки
opts = {
    "alias": ('кеша','кеш','инокентий','иннокентий','кишун','киш',
              'кишаня','кяш','кяша','кэш','кэша'),
    "tbr": ('скажи','расскажи','покажи','сколько','произнеси'),
    "cmds": {
        "ctime": ('текущее время','сейчас времени','который час'),
        "radio": ('включи музыку','воспроизведи радио','включи радио'),
        "stupid1": ('расскажи анекдот','рассмеши меня','ты знаешь анекдоты')
    }
}

# функции
def speak(what):
    print( what )
    speak_engine.say( what )
    speak_engine.runAndWait()
    speak_engine.stop()

def callback(recognizer, audio):
    try:
        voice = recognizer.recognize_google(audio, language="ru-RU").lower()
        print("[log] Распознано: " + voice)
    
        if voice.startswith(opts["alias"]):
            # обращаются к Кеше
            cmd = voice

            for x in opts['alias']:
                cmd = cmd.replace(x, "").strip()
            
            for x in opts['tbr']:
                cmd = cmd.replace(x, "").strip()
            
            # распознаем и выполняем команду
            cmd = recognize_cmd(cmd)
            execute_cmd(cmd['cmd'])

    except sr.UnknownValueError:
        print("[log] Голос не распознан!")
    except sr.RequestError as e:
        print("[log] Неизвестная ошибка, проверьте интернет!")

def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0}
    for c,v in opts['cmds'].items():

        for x in v:
            vrt = fuzz.ratio(cmd, x)
            if vrt > RC['percent']:
                RC['cmd'] = c
                RC['percent'] = vrt
    
    return RC

def execute_cmd(cmd):
    if cmd == 'ctime':
        # сказать текущее время
        now = datetime.datetime.now()
        speak("Сейчас " + str(now.hour) + ":" + str(now.minute))
    
    elif cmd == 'radio':
        # воспроизвести радио
        os.system("D:\\Jarvis\\res\\radio_record.m3u")
    
    elif cmd == 'stupid1':
        # рассказать анекдот
        speak("Мой разработчик не научил меня анекдотам ... Ха ха ха")
    
    else:
        print('Команда не распознана, повторите!')

# запуск
# r = sr.Recognizer()
# m = sr.Microphone(device_index = 1)
#
# with m as source:
#     r.adjust_for_ambient_noise(source)
#
speak_engine = pyttsx3.init()

# Только если у вас установлены голоса для синтеза речи!
# voices = speak_engine.getProperty('voices')
# speak_engine.setProperty('voice', voices[4].id)

# forced cmd test
# speak("Мой разработчик не научил меня анекдотам ... Ха ха ха")

#speak("Добрый день, повелитель")
#speak("Кеша слушает")

# stop_listening = r.listen_in_background(m, callback)
# while True:
#     time.sleep(0.1) # infinity loop



#  от PythonToday https://www.youtube.com/watch?v=ZZVWae8E9K0

class MyRecogn():
    def __init__(self):
        self.sr = speech_recognition.Recognizer()
        self.sr.pause_threshold = 0.5

        # завышенный порог шума выставляем, чтобы не было зависании на записи
        # https://stackoverflow.com/questions/32753415/python-speechrecognition-ignores-timeout-when-listening-and-hangs
        self.sr.dynamic_energy_threshold = False
        self.mes_init = 'Уровень шума до замера микрофона = %d' % self.sr.energy_threshold
        with speech_recognition.Microphone(device_index=1) as mic:
            self.sr.adjust_for_ambient_noise(source=mic)
        self.mes_init += '\nУровень шума после замера микрофона = %d' % self.sr.energy_threshold
        print(self.mes_init)
        # self.sr.energy_threshold = self.sr.energy_threshold + 500
        self.sr.energy_threshold = 2500
        # self.sr.dynamic_energy_adjustment_damping = 0.05
        # self.sr.dynamic_energy_adjustment_ratio = 1.5

    def listen_command(self, signal_status=None):
        try:
            with speech_recognition.Microphone(device_index=1) as mic:
                # signal_status.emit('Замеряем шум')
                # self.sr.adjust_for_ambient_noise(source=mic)
                # self.sr.energy_threshold = self.sr.energy_threshold + 150
                if signal_status:
                    signal_status.emit('Говорите')
                print('Говорите')
                try:
                    audio = self.sr.listen(source=mic)
                except speech_recognition.WaitTimeoutError:
                    if signal_status:
                        signal_status.emit('Четче говорите')
                    print('Четче говорите')
                    return ''
                if signal_status:
                    signal_status.emit('Отправляем в гугл')
                print('Отправляем в гугл')
                query = self.sr.recognize_google(audio_data=audio, language='ru-Ru')  #.lower()
                # query = self.sr.recognize_google(audio_data=audio, language='en-En').lower()
            print('распознано: ' + query)
            return query
        except speech_recognition.UnknownValueError:
            if signal_status:
                signal_status.emit('Не удалось распознать')
            print('Не удалось распознать')
            return ''

    def greeting(self):
        speak('Слушаю')
        query = self.listen_command()
        speak(query)

    def editing_text_dialog(self):
        pass


    def pipeReq(self, mes, attempts = 60):
        if debug_level >= 3:
            print (mes)
        for ii in range(attempts):
            try:
                pipe_req = os.open("\\\\.\\pipe\\audio_bot_answer_%d" % 401, os.O_RDWR)
                break
            except Exception as e:
                if debug_level >= 3:
                    print ("new attempt to create req pipe")
                time.sleep(1)
        try:
            pipe_req
        except Exception as e:
            print("check if ACE started!")
            print(e)
            return False
        os.write(pipe_req, mes.encode())
        os.close(pipe_req)
        return pipe_req
    def createAnswerPipe(self):
        try:
            p = win32pipe.CreateNamedPipe(r'\\.\pipe\audio_bot_answer_%d' % 369,
                                          win32pipe.PIPE_ACCESS_DUPLEX,
                                          win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
                                          1, 65536, 65536, 300, None)
        except Exception as e:
            print(e)
            p = win32pipe.CreateNamedPipe(r'\\.\pipe\audio_bot_answer_%d' % 369,
                                          win32pipe.PIPE_ACCESS_DUPLEX,
                                          win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
                                          1, 65536, 65536, 300, None)
        return p

    def getAnswer(self, pipe):
        for ii in range(10):
            try:
                if not use_overlapped:
                    rc = win32pipe.ConnectNamedPipe(pipe, None)
                else:
                    overlapped = pywintypes.OVERLAPPED()
                    overlapped.hEvent = win32event.CreateEvent(None, 0, 0, None)
                    rc = win32pipe.ConnectNamedPipe(pipe, overlapped)
                    if rc == winerror.ERROR_PIPE_CONNECTED:
                        win32event.SetEvent(overlapped.hEvent)
                    rc = win32event.WaitForSingleObject(overlapped.hEvent, 10000)
                break
            except Exception as e:
                print(e)
                if ii > 2:
                    import ipdb; ipdb.set_trace()
                return bytes("ConnectNamedPipe error", 'cp1251')
                # try:
                #     win32pipe.DisconnectNamedPipe(pipe)
                # except Exception as e:
                #     print(e)
                # time.sleep(0.5)
                # print("before getAnswer 2")
                # # p.close()
                # # del p
                # # p = self.createAnswerPipe()
                # return self.getAnswer(pipe, attempt + 1)
        full_ans = win32file.ReadFile(pipe, 4096)
        win32pipe.DisconnectNamedPipe(pipe)
        pipe.close()
        del pipe
        return full_ans[1]

def main():
    my_rec = MyRecogn()

    # p = my_rec.createAnswerPipe()
    # pipe_descr = my_rec.pipeReq('test')
    # answer = my_rec.getAnswer(p).decode('cp1251')
    # my_rec.editing_text_dialog()  # test

    while True:
        query = my_rec.listen_command()
        for k, v in my_rec.commands_dict['commands'].items():
            if query in v:
                print(k)
                globals()[k]()


# main()
