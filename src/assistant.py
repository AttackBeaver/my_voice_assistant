import os
import json
import time
import threading
import queue
import sys
import subprocess
import random
import re
import sounddevice as sd
import numpy as np
import pyautogui
import screeninfo
import requests
import ctypes
from ctypes import wintypes
from vosk import Model, KaldiRecognizer
from datetime import datetime
import win32com.client

# ===== КОНФИГУРАЦИЯ =====
MODEL_PATH = "model"
WAKE_WORD = "руна"
COOLDOWN_SECONDS = 2

YANDEX_BROWSER_PATH = None

FOLDERS_TO_OPEN = [
    r"D:\Projects",
    r"E:\Работа",
]

# ===== СЛОВАРЬ ЧИСЛИТЕЛЬНЫХ =====
NUM_WORDS = {
    'один': 1, 'одну': 1, 'одна': 1,
    'два': 2, 'две': 2,
    'три': 3,
    'четыре': 4,
    'пять': 5,
    'шесть': 6,
    'семь': 7,
    'восемь': 8,
    'девять': 9,
    'десять': 10,
    'одиннадцать': 11,
    'двенадцать': 12,
    'тринадцать': 13,
    'четырнадцать': 14,
    'пятнадцать': 15,
    'шестнадцать': 16,
    'семнадцать': 17,
    'восемнадцать': 18,
    'девятнадцать': 19,
    'двадцать': 20,
    'тридцать': 30,
    'сорок': 40,
    'пятьдесят': 50,
    'шестьдесят': 60,
}

# ========================

class Assistant:
    def __init__(self):
        # --- Единый экземпляр SAPI для синтеза ---
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        # Настройка скорости и громкости
        self.speaker.Rate = 5        # скорость: -10..10, 0 — норма, 2 — быстрее
        self.speaker.Volume = 100    # громкость: 0..100

        self.model = self._load_model()
        self.recognizer = KaldiRecognizer(self.model, 16000)
        self.recognizer.SetWords(False)

        self.sample_rate = 16000
        self.block_size = 1000
        self.audio_queue = queue.Queue()

        self.is_active = False
        self.cooldown_until = 0
        self.running = True

        # Состояние подтверждения для выключения/перезагрузки
        self.waiting_confirmation = False
        self.confirmation_action = None
        self.confirmation_timer = None

        self.speak(self._get_greeting_info())
        print("Ассистент запущен. Скажите 'Руна' для активации.")

    # ------------------ Погода и приветствие ------------------
    def _get_weather(self):
        try:
            response = requests.get("http://wttr.in/?format=%C+%t", timeout=10).text.strip()
            if not response:
                return ""

            parts = response.rsplit(' ', 1)
            if len(parts) != 2:
                return ""
            condition_en = parts[0].strip()
            temp_raw = parts[1].strip()

            weather_translation = {
                "Clear": "ясно", "Sunny": "солнечно", "Partly cloudy": "переменная облачность",
                "Cloudy": "облачно", "Overcast": "пасмурно", "Light rain": "небольшой дождь",
                "Rain": "дождь", "Heavy rain": "сильный дождь", "Snow": "снег",
                "Light snow": "небольшой снег", "Heavy snow": "сильный снег", "Fog": "туман",
                "Mist": "дымка", "Thunderstorm": "гроза", "Drizzle": "моросящий дождь",
                "Patchy rain nearby": "местами дождь",
            }
            condition_ru = condition_en
            for eng, rus in weather_translation.items():
                if eng.lower() in condition_en.lower():
                    condition_ru = rus
                    break

            temp = temp_raw.replace("°C", "градусов").strip()
            return f"{condition_ru}, {temp}"
        except Exception as e:
            print(f"Не удалось получить погоду: {e}")
            return ""

    def _get_greeting_info(self):
        now = datetime.now()
        hour = now.hour
        if 6 <= hour < 12:
            greeting = "Доброе утро"
        elif 12 <= hour < 18:
            greeting = "Добрый день"
        elif 18 <= hour < 23:
            greeting = "Добрый вечер"
        else:
            greeting = "Доброй ночи"

        time_str = now.strftime("%H:%M")
        weekdays = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]
        weekday = weekdays[now.weekday()]
        months = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"]
        date_str = f"{now.day} {months[now.month-1]}"

        weather = self._get_weather()
        weather_str = f", за окном: {weather}" if weather else ""

        return f"{greeting}! Сегодня {weekday}, {date_str}, местное время {time_str}{weather_str}."

    # ------------------ Модель распознавания ------------------
    def _load_model(self):
        if not os.path.exists(MODEL_PATH):
            raise FileNotFoundError(f"Папка модели '{MODEL_PATH}' не найдена.")
        items = os.listdir(MODEL_PATH)
        for item in items:
            full_path = os.path.join(MODEL_PATH, item)
            if os.path.isdir(full_path) and "vosk" in item.lower():
                print(f"Загружаем модель из: {full_path}")
                return Model(full_path)
        raise FileNotFoundError("Не найдена папка с моделью Vosk внутри 'model/'.")

    # ------------------ Синтез речи (SAPI) с прерыванием ------------------
    def speak(self, text):
        print(f"Руна: {text}")
        # Останавливаем текущую речь через очистку очереди
        self.speaker.Speak("", 2)  # 2 = SVSFPurgeBeforeSpeak (без воспроизведения)
        # Запускаем новую речь асинхронно с очисткой
        self.speaker.Speak(text, 1)  # 1 = SVSFlagsAsync | SVSFPurgeBeforeSpeak

    def audio_callback(self, indata, frames, time, status):
        if status:
            print(f"Статус аудио: {status}")
        self.audio_queue.put(bytes(indata))

    # ------------------ Вспомогательные функции для окон ------------------
    def _find_yandex_browser(self):
        possible_paths = [
            os.path.expanduser(r"~\AppData\Local\Yandex\YandexBrowser\Application\browser.exe"),
            r"C:\Program Files (x86)\Yandex\YandexBrowser\Application\browser.exe",
            r"C:\Program Files\Yandex\YandexBrowser\Application\browser.exe",
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None

    def _open_folders(self):
        for folder in FOLDERS_TO_OPEN:
            if os.path.exists(folder):
                subprocess.Popen(['explorer', folder])
                time.sleep(0.5)
            else:
                print(f"Папка не найдена: {folder}")

    def _open_browser_windows(self):
        browser_path = YANDEX_BROWSER_PATH or self._find_yandex_browser()
        if not browser_path or not os.path.exists(browser_path):
            self.speak("Яндекс.Браузер не найден. Проверьте путь.")
            return

        monitors = screeninfo.get_monitors()
        if len(monitors) < 2:
            self.speak("Обнаружено меньше двух мониторов. Открываю одно окно.")
            subprocess.Popen([browser_path])
            time.sleep(2)
            pyautogui.hotkey('win', 'up')
            return

        monitors_sorted = sorted(monitors, key=lambda m: m.x)
        primary = next((m for m in monitors_sorted if m.is_primary), monitors_sorted[0])
        secondary = next((m for m in monitors_sorted if not m.is_primary), None)

        self.speak("Открываю рабочие окна.")
        proc1 = subprocess.Popen([browser_path])
        time.sleep(1.5)
        proc2 = subprocess.Popen([browser_path])
        time.sleep(1.5)

        windows = pyautogui.getWindowsWithTitle('Яндекс')
        if not windows:
            windows = pyautogui.getWindowsWithTitle('Yandex')
        if not windows:
            self.speak("Не удалось найти окна браузера. Попробуйте вручную.")
            return

        browser_windows = windows[-2:] if len(windows) >= 2 else windows

        if len(browser_windows) >= 2:
            win1 = browser_windows[0]
            win1.restore()
            win1.moveTo(secondary.x, secondary.y)
            win1.resizeTo(secondary.width, secondary.height)
            win1.maximize()

            win2 = browser_windows[1]
            win2.restore()
            win2.moveTo(primary.x, primary.y)
            win2.resizeTo(primary.width, primary.height)
            win2.maximize()
        else:
            browser_windows[0].restore()
            browser_windows[0].moveTo(primary.x, primary.y)
            browser_windows[0].resizeTo(primary.width, primary.height)
            browser_windows[0].maximize()

        time.sleep(1)

    def _open_browser_url(self, url):
        try:
            subprocess.Popen(['start', url], shell=True)
            return True
        except Exception as e:
            print(f"Ошибка открытия URL: {e}")
            return False

    def _open_browser_on_secondary_monitor(self, url):
        try:
            monitors = screeninfo.get_monitors()
            if len(monitors) < 2:
                self.speak("Обнаружено меньше двух мониторов. Открываю на основном.")
                self._open_browser_url(url)
                return

            monitors_sorted = sorted(monitors, key=lambda m: m.x)
            secondary = monitors_sorted[0]

            subprocess.Popen(['start', url], shell=True)
            time.sleep(2)

            windows = pyautogui.getWindowsWithTitle('Яндекс')
            if not windows:
                windows = pyautogui.getWindowsWithTitle('Yandex')
            if not windows:
                windows = pyautogui.getWindowsWithTitle('Chrome')
            if not windows:
                windows = pyautogui.getWindowsWithTitle('Firefox')
            if not windows:
                self.speak("Не удалось найти окно браузера. Попробуйте вручную.")
                return

            win = windows[-1]
            win.restore()
            win.moveTo(secondary.x, secondary.y)
            win.resizeTo(secondary.width, secondary.height)
            win.maximize()
            print("Окно браузера перемещено на второй монитор.")
        except Exception as e:
            print(f"Ошибка открытия на втором мониторе: {e}")
            self.speak("Не удалось открыть браузер на втором мониторе.")

    # ------------------ Надёжные функции управления окнами ------------------
    def _minimize_all_windows(self):
        try:
            user32 = ctypes.windll.user32
            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
            def enum_callback(hwnd, lParam):
                if user32.IsWindowVisible(hwnd):
                    user32.ShowWindow(hwnd, 6)
                return True
            user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
            print("Сворачивание через WinAPI выполнено")
            return True
        except Exception as e:
            print(f"Ошибка сворачивания через WinAPI: {e}")
            return False

    def _restore_all_windows(self):
        try:
            user32 = ctypes.windll.user32
            def enum_callback(hwnd, lParam):
                if user32.IsWindowVisible(hwnd) and user32.IsIconic(hwnd):
                    user32.ShowWindow(hwnd, 9)
                return True
            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
            user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
            print("Разворачивание через WinAPI выполнено")
            return True
        except Exception as e:
            print(f"Ошибка разворачивания через WinAPI: {e}")
            return False

    # ------------------ Таймер-напоминалка ------------------
    def _parse_time_amount(self, text):
        # Сначала ищем цифру
        match = re.search(r'(\d+)\s+(минут|минуту|минуты|секунд|секунду|секунды|час|часа|часов)', text)
        if match:
            amount = int(match.group(1))
            unit = match.group(2)
            return amount, unit, match.start(), match.end()
        # Ищем слово из словаря
        for word, num in NUM_WORDS.items():
            pattern = rf'\b{word}\s+(минут|минуту|минуты|секунд|секунду|секунды|час|часа|часов)\b'
            match = re.search(pattern, text)
            if match:
                amount = num
                unit = match.group(1)
                return amount, unit, match.start(), match.end()
        return None, None, None, None

    def _set_reminder(self, text):
        amount, unit, start, end = self._parse_time_amount(text)
        if amount is None:
            self.speak("Не поняла время. Скажите, например: 'напомни через 5 минут позвонить'.")
            return

        # Преобразуем единицу в секунды
        if unit.startswith('минут'):
            delay = amount * 60
        elif unit.startswith('секунд'):
            delay = amount
        elif unit.startswith('час'):
            delay = amount * 3600
        else:
            self.speak("Неизвестная единица времени.")
            return

        # Извлекаем текст напоминания: удаляем найденную временную часть и слово "напомни"
        reminder_text = text[:start] + text[end:]
        # Удаляем слово "напомни" (и варианты) и лишние пробелы
        for word in ["напомни", "напомнить", "напомни через"]:
            reminder_text = reminder_text.replace(word, "").strip()
        # Убираем "через" если осталось
        reminder_text = reminder_text.replace("через", "").strip()
        if not reminder_text:
            reminder_text = "напоминание"

        self.speak(f"Хорошо, напомню через {amount} {unit}.")

        def reminder_job():
            time.sleep(delay)
            self.speak(f"Напоминание: {reminder_text}")

        threading.Thread(target=reminder_job, daemon=True).start()

    # ------------------ Подтверждение выключения/перезагрузки ------------------
    def _start_confirmation(self, action):
        self.waiting_confirmation = True
        self.confirmation_action = action
        self.speak("Вы уверены? Скажите 'да' или 'нет'.")
        def timeout_cancel():
            time.sleep(10)
            if self.waiting_confirmation:
                self.waiting_confirmation = False
                self.confirmation_action = None
                self.speak("Время вышло. Действие отменено.")
        self.confirmation_timer = threading.Thread(target=timeout_cancel, daemon=True)
        self.confirmation_timer.start()

    def _process_confirmation(self, text):
        if "да" in text:
            self.waiting_confirmation = False
            if self.confirmation_timer:
                self.confirmation_timer.join(0)
            if self.confirmation_action == "shutdown":
                self.speak("Выключаю компьютер.")
                os.system("shutdown /s /t 0")
            elif self.confirmation_action == "restart":
                self.speak("Перезагружаю компьютер.")
                os.system("shutdown /r /t 0")
            self.confirmation_action = None
        elif "нет" in text:
            self.waiting_confirmation = False
            if self.confirmation_timer:
                self.confirmation_timer.join(0)
            self.speak("Действие отменено.")
            self.confirmation_action = None
        else:
            self.speak("Я не поняла. Скажите 'да' или 'нет'.")

    # ------------------ Обработка команд ------------------
    def process_command(self, text):
        text_lower = text.lower().strip()
        # Удаляем служебные фразы (все в нижнем регистре)
        for phrase in ["руна", "слушаем вас", "слушаю вас", "жаль вас з", "служит вас"]:
            text_lower = text_lower.replace(phrase, "").strip()
        print(f"Распознано (команда после очистки): '{text_lower}'")

        # Выход
        if "стоп" in text_lower or "выход" in text_lower:
            self.speak("До свидания!")
            self.running = False
            return

        # Похвала
        praise_phrases = ["спасибо", "ты молодец", "умница", "отлично", "супер", "классно", "круто", "хорошая работа", "молодец"]
        if any(phrase in text_lower for phrase in praise_phrases):
            responses = [
                "Спасибо, приятно слышать.",
                "Рада стараться.",
                "Всегда к вашим услугам.",
                "Очень приятно.",
                "Стараюсь!"
            ]
            self.speak(random.choice(responses))
            self.start_cooldown()
            return

        # Справка
        if any(phrase in text_lower for phrase in ["что умеешь", "расскажи о себе", "твои возможности", "помощь", "справка"]):
            self.speak(
                "Я умею завершать работу по команде 'стоп' или 'выход'. "
                "Могу рассказать о своих возможностях. "
                "Также я умею открывать рабочие окна по команде 'работаем'. "
                "Могу закрыть все окна, свернуть или развернуть все окна, показать рабочий стол. "
                "Ещё я знаю, как дела, умею танцевать, показывать погоду, напоминать о делах, искать в интернете, "
                "а также выключать или перезагружать компьютер."
            )
            self.start_cooldown()
            return

        # Рабочий режим
        if any(phrase in text_lower for phrase in ["работаем", "работа", "рабочий режим", "воркать", "на работу"]):
            self.speak("Запускаю рабочий режим.")
            self._open_folders()
            self._open_browser_windows()
            self.speak("Рабочий режим активирован.")
            self.start_cooldown()
            return

        # Управление окнами
        if "закр" in text_lower:
            if "активное окно" in text_lower or "активную" in text_lower:
                self.speak("Закрываю активное окно.")
                pyautogui.hotkey('alt', 'f4')
                self.speak("Готово.")
                self.start_cooldown()
                return
            else:
                self.speak("Закрываю все окна.")
                for _ in range(5):
                    pyautogui.hotkey('alt', 'f4')
                    time.sleep(0.3)
                self.speak("Все окна закрыты.")
                self.start_cooldown()
                return

        if "сверн" in text_lower or "шухер" in text_lower or "шмон" in text_lower:
            self.speak("Сворачиваю все окна.")
            if self._minimize_all_windows():
                self.speak("Готово.")
            else:
                self.speak("Не удалось свернуть окна.")
            self.start_cooldown()
            return

        if "разверн" in text_lower:
            self.speak("Разворачиваю все окна.")
            if self._restore_all_windows():
                self.speak("Готово.")
            else:
                self.speak("Не удалось развернуть окна.")
            self.start_cooldown()
            return

        if "рабочий стол" in text_lower or "покажи стол" in text_lower:
            self.speak("Показываю рабочий стол.")
            pyautogui.hotkey('win', 'd')
            self.speak("Готово.")
            self.start_cooldown()
            return

        # Новые команды
        if "как дела" in text_lower or "как твои дела" in text_lower:
            answers = [
                "Хорошо, спасибо что спросили!",
                "Отлично! А у вас?",
                "Всё замечательно, работаю.",
                "Прекрасно, я всегда готова помочь!"
            ]
            self.speak(random.choice(answers))
            self.start_cooldown()
            return

        if "танцуем" in text_lower or "танцы" in text_lower:
            self.speak("Танцуют все!")
            self._open_browser_url("https://youtu.be/dQw4w9WgXcQ?si=vCrGntBRfPiwPciy&t=43")
            self.start_cooldown()
            return

        if "выключи" in text_lower or "выключить" in text_lower:
            self._start_confirmation("shutdown")
            return
        if "перезагрузи" in text_lower or "перезагрузка" in text_lower:
            self._start_confirmation("restart")
            return

        if "поздоровайся" in text_lower or "скажи привет" in text_lower:
            self.speak(self._get_greeting_info())
            self.start_cooldown()
            return

        # Напоминание
        if "напомни" in text_lower:
            self._set_reminder(text_lower)
            self.start_cooldown()
            return

        if "погода" in text_lower:
            weather = self._get_weather()
            if weather:
                self.speak(f"За окном: {weather}")
            else:
                self.speak("Не удалось получить погоду.")
            self.start_cooldown()
            return

        # Поиск в интернете
        if "найди в интернете" in text_lower or "поищи в интернете" in text_lower:
            if "втором мониторе" in text_lower or "на втором мониторе" in text_lower or "на второй монитор" in text_lower:
                match = re.search(r'(найди в интернете|поищи в интернете)\s*(.+)', text_lower)
                if match:
                    query = match.group(2).strip()
                    for phrase in ["втором мониторе", "на втором мониторе", "на второй монитор"]:
                        query = query.replace(phrase, "").strip()
                    if query:
                        self.speak(f"Ищу в интернете на втором мониторе: {query}")
                        search_url = f"https://yandex.ru/search/?text={query.replace(' ', '+')}"
                        self._open_browser_on_secondary_monitor(search_url)
                    else:
                        self.speak("Что именно искать? Скажите, например: 'найди в интернете на втором мониторе рецепт борща'.")
                else:
                    self.speak("Не поняла запрос.")
            else:
                match = re.search(r'(найди в интернете|поищи в интернете)\s*(.+)', text_lower)
                if match:
                    query = match.group(2).strip()
                    if query:
                        self.speak(f"Ищу в интернете: {query}")
                        search_url = f"https://yandex.ru/search/?text={query.replace(' ', '+')}"
                        self._open_browser_url(search_url)
                    else:
                        self.speak("Что именно искать? Скажите, например: 'найди в интернете рецепт борща'.")
                else:
                    self.speak("Не поняла запрос.")
            self.start_cooldown()
            return

        self.speak("Не поняла вас.")
        self.start_cooldown()

    def start_cooldown(self):
        self.cooldown_until = time.time() + COOLDOWN_SECONDS
        self.is_active = False
        print(f"Кулдаун {COOLDOWN_SECONDS} сек. Игнорирую 'Руна' до {time.ctime(self.cooldown_until)}")

    # ------------------ Главный цикл ------------------
    def listen_loop(self):
        with sd.RawInputStream(samplerate=self.sample_rate, channels=1, dtype='int16',
                               blocksize=self.block_size, callback=self.audio_callback):
            while self.running:
                data = self.audio_queue.get()
                if self.recognizer.AcceptWaveform(data):
                    result = json.loads(self.recognizer.Result())
                    text = result.get("text", "").strip().lower()
                    if not text:
                        continue

                    print(f"Услышал (финал): {text}")

                    now = time.time()
                    if self.waiting_confirmation:
                        self._process_confirmation(text)
                        self.recognizer.Reset()
                        continue

                    if now < self.cooldown_until:
                        print(f"На кулдауне (осталось {int(self.cooldown_until - now)} сек). Игнорирую.")
                        self.recognizer.Reset()
                        continue

                    if not self.is_active:
                        if WAKE_WORD in text:
                            self.is_active = True
                            self.speak("Да-да.")
                            self.recognizer.Reset()
                        continue

                    cmd_text = text.replace(WAKE_WORD, "").strip()
                    if not cmd_text:
                        continue

                    self.process_command(cmd_text)
                    self.is_active = False
                    self.recognizer.Reset()

                else:
                    partial = json.loads(self.recognizer.PartialResult())
                    partial_text = partial.get("partial", "").strip().lower()
                    if partial_text:
                        now = time.time()
                        if not self.waiting_confirmation and now >= self.cooldown_until and not self.is_active:
                            if WAKE_WORD in partial_text:
                                self.is_active = True
                                self.speak("Да-да.")
                                # Не сбрасываем recognizer, чтобы накопить команду

if __name__ == "__main__":
    assistant = Assistant()
    try:
        assistant.listen_loop()
    except KeyboardInterrupt:
        print("Завершено пользователем.")
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        assistant.speak("Произошла ошибка. Перезапустите ассистента.")