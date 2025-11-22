# -*- coding: utf-8 -*-
import os
import sys
import time
import ssl
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import win32print
from imap_tools import MailBox, A
import re
import traceback
import configparser
import multiprocessing

# --- НАСТРОЙКИ ПО УМОЛЧАНИЮ ---
DEFAULT_IMAP_SERVER = "imap.ultramed-nn.ru"
DEFAULT_IMAP_USER = "print@ultramed-nn.ru"
PROCESSED_FOLDER_NAME = "Printed"
TEMP_FOLDER_NAME = "Temp_Print"
DEFAULT_IRFANVIEW_PATH = r"C:\Program Files\IrfanView\i_view64.exe"
DEFAULT_SUMATRA_PATH = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
IMAGE_PRINTERS = ["IrfanView (Рекомендуемый)", "MS Paint"]
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
PDF_EXTENSION = '.pdf'
SETTINGS_FILE_NAME = "settings.ini"

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def mark_all_as_processed_thread(settings, log_queue):
    log_queue.put("--- Запущена очистка очереди... ---")
    if not all(settings[k] for k in ["server", "user", "password"]):
        log_queue.put("!!! ОШИБКА: Для очистки очереди необходимо заполнить данные для входа.")
        return
    try:
        ssl._create_default_https_context = ssl._create_unverified_context
        with MailBox(settings["server"]).login(settings["user"], settings["password"], 'INBOX') as mailbox:
            if not mailbox.folder.exists(PROCESSED_FOLDER_NAME):
                log_queue.put(f"!!! ОШИБКА: Папка '{PROCESSED_FOLDER_NAME}' не найдена!")
                return
            uids = mailbox.uids(A(seen=False))
            if not uids:
                log_queue.put("-> Непрочитанных писем для очистки нет.")
                return
            log_queue.put(f"-> Найдено {len(uids)} писем. Перемещаю в '{PROCESSED_FOLDER_NAME}'...")
            mailbox.move(uids, PROCESSED_FOLDER_NAME)
            log_queue.put(f"=== Очистка очереди завершена. {len(uids)} писем перемещено. ===")
    except Exception as e:
        tb_str = traceback.format_exc()
        log_queue.put(f"!!! КРИТИЧЕСКАЯ ОШИБКА при очистке: {e}")
        log_queue.put(f"!!! TRACEBACK:\n{tb_str}")

def safe_string(s): return ''.join(c for c in s if c.isprintable()) if s else ""


class PrintBotApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Печать с Почты v5.0 (Persistent Connection)")
        self.geometry("750x800")
        self.worker_thread = None
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        self.base_dir = get_base_path()
        self.settings_path = os.path.join(self.base_dir, SETTINGS_FILE_NAME)
        self.temp_dir = os.path.join(self.base_dir, TEMP_FOLDER_NAME)
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        self.create_widgets()
        self.process_log_queue()
        self.load_settings()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Connection'] = {'server': self.entry_server.get(),'user': self.entry_user.get(),'password': self.entry_password.get()}
        config['Printing'] = {
            'image_printer_device': self.image_printer_device_var.get(),
            'pdf_printer_device': self.pdf_printer_device_var.get(),
            'image_handler_program': self.image_handler_program_var.get(),
            'irfanview_path': self.entry_irfanview_path.get(),
            'sumatra_path': self.entry_pdf_printer_path.get()
        }
        config['Filtering'] = {'print_all': str(self.print_all_var.get()),'whitelist': self.whitelist_text.get("1.0", tk.END).strip()}
        try:
            with open(self.settings_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            self.log("-> Настройки успешно сохранены.")
        except Exception as e:
            self.log(f"!!! Ошибка сохранения настроек: {e}")

    def load_settings(self):
        config = configparser.ConfigParser()
        if not os.path.exists(self.settings_path):
            self.log(f"-> Файл настроек не найден. Будет создан при выходе.")
            return
        try:
            config.read(self.settings_path, encoding='utf-8')
            default_printer = win32print.GetDefaultPrinter()
            self.entry_server.set(config.get('Connection', 'server', fallback=DEFAULT_IMAP_SERVER))
            self.entry_user.set(config.get('Connection', 'user', fallback=DEFAULT_IMAP_USER))
            self.entry_password.set(config.get('Connection', 'password', fallback=''))
            self.image_printer_device_var.set(config.get('Printing', 'image_printer_device', fallback=default_printer))
            self.pdf_printer_device_var.set(config.get('Printing', 'pdf_printer_device', fallback=default_printer))
            self.image_handler_program_var.set(config.get('Printing', 'image_handler_program', fallback=IMAGE_PRINTERS[0]))
            self.entry_irfanview_path.set(config.get('Printing', 'irfanview_path', fallback=DEFAULT_IRFANVIEW_PATH))
            self.entry_pdf_printer_path.set(config.get('Printing', 'sumatra_path', fallback=DEFAULT_SUMATRA_PATH))
            self.print_all_var.set(config.getboolean('Filtering', 'print_all', fallback=True))
            self.whitelist_text.delete("1.0", tk.END)
            self.whitelist_text.insert("1.0", config.get('Filtering', 'whitelist', fallback=''))
            self.toggle_whitelist_state()
            self.log("-> Настройки успешно загружены.")
        except Exception as e:
            self.log(f"!!! Ошибка загрузки настроек: {e}")
            messagebox.showerror("Ошибка загрузки настроек", f"Не удалось прочитать файл settings.ini.\n\nОшибка: {e}")

    def on_closing(self):
        if messagebox.askokcancel("Выход", "Сохранить настройки и выйти?"):
            self.save_settings()
            self.destroy()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки подключения", padding="10")
        settings_frame.pack(fill=tk.X, pady=5)
        settings_frame.columnconfigure(1, weight=1)
        self.entry_server = tk.StringVar(value=DEFAULT_IMAP_SERVER)
        ttk.Label(settings_frame, text="IMAP Сервер:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.entry_server).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.entry_user = tk.StringVar(value=DEFAULT_IMAP_USER)
        ttk.Label(settings_frame, text="Email:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.entry_user).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.entry_password = tk.StringVar()
        ttk.Label(settings_frame, text="Пароль:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.entry_password, show="*").grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        print_settings_frame = ttk.LabelFrame(main_frame, text="Настройки печати", padding="10")
        print_settings_frame.pack(fill=tk.X, pady=5)
        print_settings_frame.columnconfigure(1, weight=1)
        printers = [p[2] for p in win32print.EnumPrinters(2)]
        default_printer = win32print.GetDefaultPrinter()
        ttk.Label(print_settings_frame, text="Печать картинок через:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.image_handler_program_var = tk.StringVar(value=IMAGE_PRINTERS[0])
        ttk.Combobox(print_settings_frame, textvariable=self.image_handler_program_var, values=IMAGE_PRINTERS, state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(print_settings_frame, text="Принтер для картинок:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.image_printer_device_var = tk.StringVar(value=default_printer)
        ttk.Combobox(print_settings_frame, textvariable=self.image_printer_device_var, values=printers, state="readonly").grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.entry_irfanview_path = self.create_path_picker(print_settings_frame, "Путь к IrfanView:", 2, DEFAULT_IRFANVIEW_PATH)
        ttk.Separator(print_settings_frame, orient='horizontal').grid(row=3, columnspan=3, sticky='ew', pady=10)
        ttk.Label(print_settings_frame, text="Принтер для PDF:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.pdf_printer_device_var = tk.StringVar(value=default_printer)
        ttk.Combobox(print_settings_frame, textvariable=self.pdf_printer_device_var, values=printers, state="readonly").grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        self.entry_pdf_printer_path = self.create_path_picker(print_settings_frame, "Путь к SumatraPDF:", 5, DEFAULT_SUMATRA_PATH)
        whitelist_frame = ttk.LabelFrame(main_frame, text="Фильтр отправителей", padding="10")
        whitelist_frame.pack(fill=tk.BOTH, pady=5, expand=True)
        self.print_all_var = tk.BooleanVar(value=True)
        self.print_all_check = ttk.Checkbutton(whitelist_frame, text="Печатать отовсюду (игнорировать белый список)", variable=self.print_all_var,command=self.toggle_whitelist_state)
        self.print_all_check.pack(anchor="w")
        ttk.Label(whitelist_frame, text="Белый список (email через запятую или с новой строки):").pack(anchor="w", pady=(5,0))
        self.whitelist_text = tk.Text(whitelist_frame, height=4, font=("Consolas", 9))
        self.whitelist_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.toggle_whitelist_state()
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(padx=10, pady=5, fill="x", expand=False)
        self.start_button = ttk.Button(control_frame, text="Старт", command=self.start_worker)
        self.start_button.pack(side=tk.LEFT, padx=5, fill="x", expand=True)
        self.stop_button = ttk.Button(control_frame, text="Стоп", command=self.stop_worker, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5, fill="x", expand=True)
        self.mark_all_button = ttk.Button(control_frame, text="Очистить очередь", command=self.clear_print_queue)
        self.mark_all_button.pack(side=tk.LEFT, padx=5, fill="x", expand=True)
        self.clear_button = ttk.Button(control_frame, text=f"Очистить папку {TEMP_FOLDER_NAME}", command=self.clear_temp_folder)
        self.clear_button.pack(side=tk.LEFT, padx=5, fill="x", expand=True)
        log_frame = ttk.LabelFrame(main_frame, text="Логи", padding="10")
        log_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.log_text = tk.Text(log_frame, state="disabled", height=10, wrap="word", font=("Consolas", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text['yscrollcommand'] = scrollbar.set

    def create_path_picker(self, p, l, r, d):
        ttk.Label(p, text=l).grid(row=r, column=0, sticky="w", padx=5, pady=5)
        var = tk.StringVar(value=d)
        e=ttk.Entry(p, textvariable=var); e.grid(row=r, column=1, sticky="ew", padx=5, pady=5)
        b=ttk.Button(p, text="...", width=4, command=lambda v=var: self.browse_file(v)); b.grid(row=r, column=2, padx=5, pady=5)
        return var
    
    def toggle_whitelist_state(self):
        state = "disabled" if self.print_all_var.get() else "normal"
        bg_color = "#f0f0f0" if self.print_all_var.get() else "#ffffff"
        self.whitelist_text.config(state=state, background=bg_color)

    def browse_file(self, v): 
        f=filedialog.askopenfilename(filetypes=[("Программы", "*.exe")]);
        if f: v.set(f)

    def log(self, m): 
        self.log_queue.put(m)

    def process_log_queue(self):
        try:
            while True:
                m=self.log_queue.get_nowait()
                if m=="WORKER_STOPPED_ERROR": 
                    self.start_button.config(state="normal")
                    self.stop_button.config(state="disabled")
                    self.mark_all_button.config(state="normal")
                    continue
                self.log_text.config(state="normal")
                self.log_text.insert("end", f"{m}\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except queue.Empty: 
            pass
        self.after(200, self.process_log_queue)

    def get_settings(self):
        raw_whitelist = self.whitelist_text.get("1.0", tk.END).strip()
        cleaned_whitelist = [email.strip().lower() for email in re.split(r'[,\n]', raw_whitelist) if email.strip()]
        return {"server": self.entry_server.get().strip(),"user": self.entry_user.get().strip(),"password": self.entry_password.get(),"temp_dir": self.temp_dir,"image_printer_device": self.image_printer_device_var.get(),"pdf_printer_device": self.pdf_printer_device_var.get(),"image_handler_program": self.image_handler_program_var.get(), "irfanview_path": self.entry_irfanview_path.get(),"pdf_printer_path": self.entry_pdf_printer_path.get(),"print_all": self.print_all_var.get(),"whitelist": cleaned_whitelist}
    
    def start_worker(self):
        settings = self.get_settings()
        if not all(settings[k] for k in ["server", "user", "password"]): 
            messagebox.showerror("Ошибка", "Пожалуйста, заполните поля 'IMAP Сервер', 'Email' и 'Пароль' перед запуском.")
            return
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.mark_all_button.config(state="disabled")
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=email_worker_thread, args=(settings, self.log_queue, self.stop_event), daemon=True)
        self.worker_thread.start()
        self.log("=== Воркер запущен ===")

    def stop_worker(self):
        if self.worker_thread and self.worker_thread.is_alive(): 
            self.stop_event.set()
            self.log("--- Сигнал остановки...")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.mark_all_button.config(state="normal")
        
    def clear_print_queue(self):
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите очистить очередь?\nВсе непрочитанные письма будут перемещены в папку 'Printed' без печати."):
            threading.Thread(target=mark_all_as_processed_thread, args=(self.get_settings(), self.log_queue), daemon=True).start()
    
    def clear_temp_folder(self):
        files = os.listdir(self.temp_dir)
        if not files: 
            self.log(f"-> Папка '{TEMP_FOLDER_NAME}' пуста.")
            return
        if messagebox.askyesno("Подтверждение", f"Удалить {len(files)} файлов из папки '{TEMP_FOLDER_NAME}'?"):
            d_c, e_c = 0, 0
            for f in files:
                try: 
                    os.remove(os.path.join(self.temp_dir, f))
                    d_c += 1
                except Exception as e: 
                    self.log(f"!!! Не удалось удалить {f}: {e}")
                    e_c += 1
            self.log(f"=== Очистка завершена. Удалено: {d_c}. Ошибок: {e_c}. ===")

# --- НОВЫЙ СТАБИЛЬНЫЙ ВОРКЕР (Keep-Alive) ---
def email_worker_thread(settings, log_queue, stop_event):
    ssl._create_default_https_context = ssl._create_unverified_context
    
    # Внешний цикл - отвечает за ПЕРЕПОДКЛЮЧЕНИЕ при обрыве
    while not stop_event.is_set():
        try:
            log_queue.put("-> Подключаюсь к серверу...")
            
            # Мы открываем соединение ОДИН раз и держим его
            with MailBox(settings["server"]).login(settings["user"], settings["password"], 'INBOX') as mailbox:
                log_queue.put("-> Успешное подключение! Вхожу в режим ожидания писем.")
                
                # Проверяем папку при старте сессии
                if not mailbox.folder.exists(PROCESSED_FOLDER_NAME):
                    try:
                        mailbox.folder.create(PROCESSED_FOLDER_NAME)
                        log_queue.put(f"-> Папка '{PROCESSED_FOLDER_NAME}' создана.")
                    except Exception as e:
                        log_queue.put(f"!!! ОШИБКА: Нет папки '{PROCESSED_FOLDER_NAME}' и не могу создать: {e}")
                        log_queue.put("WORKER_STOPPED_ERROR")
                        return

                # Внутренний цикл - работает пока есть соединение
                while not stop_event.is_set():
                    # 1. Читаем письма
                    try:
                        messages_to_process = list(mailbox.fetch(A(seen=False), limit=10, reverse=True))
                        
                        if messages_to_process:
                            log_queue.put(f"-> Найдено новых писем: {len(messages_to_process)}")
                            
                            for msg in messages_to_process:
                                if stop_event.is_set(): break
                                try:
                                    sender_email = msg.from_.lower() if msg.from_ else ""
                                    log_queue.put(f"--- Получено письмо от: {sender_email} ---")
                                    
                                    # Фильтр
                                    if not settings["print_all"] and sender_email not in settings["whitelist"]:
                                        log_queue.put(f"-> Не в белом списке.")
                                        mailbox.move(msg.uid, PROCESSED_FOLDER_NAME)
                                        continue
                                    
                                    # Вложения
                                    if not msg.attachments:
                                        log_queue.put(f"-> Нет вложений.")
                                        mailbox.move(msg.uid, PROCESSED_FOLDER_NAME)
                                        continue
                                    
                                    for att in msg.attachments:
                                        safe_filename = safe_string(att.filename)
                                        filepath = os.path.join(settings["temp_dir"], safe_filename)
                                        with open(filepath, 'wb') as f: f.write(att.payload)
                                        log_queue.put(f"-> Скачан: {safe_filename}")
                                        
                                        # Пауза для записи файла на диск
                                        time.sleep(2)

                                        # ПЕЧАТЬ
                                        file_ext = os.path.splitext(safe_filename)[1].lower()
                                        if file_ext == PDF_EXTENSION:
                                            cmd = [settings["pdf_printer_path"], '-print-to', settings["pdf_printer_device"], filepath]
                                            subprocess.run(cmd, check=True, timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
                                            log_queue.put("-> PDF отправлен на печать.")
                                        
                                        elif file_ext in IMAGE_EXTENSIONS:
                                            if settings["image_handler_program"] == IMAGE_PRINTERS[0]: # IrfanView
                                                photo_printer = settings["image_printer_device"].strip()
                                                orig_printer = None
                                                try:
                                                    orig_printer = win32print.GetDefaultPrinter()
                                                    win32print.SetDefaultPrinter(photo_printer)
                                                    cmd = [settings["irfanview_path"], filepath, "/print"]
                                                    subprocess.run(cmd, check=True, timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
                                                    log_queue.put("-> Картинка отправлена (IrfanView).")
                                                finally:
                                                    if orig_printer: win32print.SetDefaultPrinter(orig_printer)
                                            else: # Paint
                                                cmd = ['mspaint.exe', '/pt', filepath, settings["image_printer_device"]]
                                                subprocess.run(cmd, check=True, timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
                                                log_queue.put("-> Картинка отправлена (Paint).")
                                        else:
                                            log_queue.put(f"-> Тип {file_ext} не поддерживается.")
                                    
                                    mailbox.move(msg.uid, PROCESSED_FOLDER_NAME)
                                    log_queue.put("--- Письмо обработано. ---")

                                except Exception as e:
                                    log_queue.put(f"!!! Ошибка обработки письма UID {msg.uid}: {e}")
                                    # Пытаемся убрать письмо, чтобы не застрять
                                    try: mailbox.move(msg.uid, PROCESSED_FOLDER_NAME)
                                    except: pass

                        # 2. Пауза (Keep-Alive "пинг")
                        # Вместо разрыва соединения, мы просто ждем 15 секунд внутри сессии
                        for _ in range(15): 
                            if stop_event.is_set(): break
                            time.sleep(1)
                            
                    except Exception as fetch_error:
                        # Если ошибка случилась при fetch (например, сокет закрылся сервером)
                        # Мы выбрасываем её выше, чтобы внешний цикл сделал переподключение
                        raise fetch_error

        except Exception as e:
            # Сюда мы попадаем, если соединение разорвалось (10060, Socket closed и т.д.)
            log_queue.put(f"!!! Связь потеряна: {e}")
            log_queue.put("-> Кулдаун 15 секунд и переподключаюсь...") 
            
            for _ in range(15):
                if stop_event.is_set(): break
                time.sleep(1)
            
    log_queue.put("=== Воркер остановлен. ===")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = PrintBotApp()
    app.mainloop()