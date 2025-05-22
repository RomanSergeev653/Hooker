import pandas as pd
import requests
import tkinter as tk
from tkinter import simpledialog, filedialog, ttk, messagebox, Listbox, Button, Frame
import time
import threading
import json
import os

# Конфигурация стиля
CONFIG_FILE = "file_history.json"
MAX_HISTORY = 5


def apply_style(root):
    """Применяет единый стиль ко всем элементам интерфейса"""
    root.option_add('*Font', 'Arial 14')
    root.option_add('*Background', 'white')
    root.option_add('*Foreground', 'black')
    root.option_add('*Button.Background', 'blue')
    root.option_add('*Button.Foreground', 'white')
    root.option_add('*Listbox.Background', 'white')
    root.option_add('*Listbox.Foreground', 'black')
    root.option_add('*Listbox.selectBackground', 'navy')
    root.option_add('*Label.Background', 'white')


def load_file_history():
    """Загружает историю файлов с сохранением стиля"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            data = json.load(f)
            return data.get("history", [])
    return []


def save_file_history(file_path):
    """Сохраняет историю с проверкой стиля"""
    if not file_path or not os.path.exists(file_path):
        return

    history = load_file_history()
    if file_path in history:
        history.remove(file_path)
    history.insert(0, file_path)
    history = history[:MAX_HISTORY]

    with open(CONFIG_FILE, "w") as f:
        json.dump({
            "history": history,
            "last_path": os.path.dirname(file_path)
        }, f)


def create_history_window():
    """Создаёт окно истории с сохранением стиля"""
    root = tk.Tk()
    apply_style(root)
    root.title("Выберите файл - История")
    root.geometry("700x400")

    # Основной фрейм
    main_frame = Frame(root, bg='white')
    main_frame.pack(pady=20, padx=20, fill='both', expand=True)

    # Заголовок
    label = tk.Label(main_frame,
                     text="Последние открытые файлы:",
                     bg='white')
    label.pack(anchor='w', pady=(0, 10))

    # Список файлов с полосой прокрутки
    scrollbar = tk.Scrollbar(main_frame)
    listbox = Listbox(main_frame,
                      yscrollcommand=scrollbar.set,
                      height=8,
                      selectbackground='navy',
                      selectforeground='white')
    scrollbar.config(command=listbox.yview)

    listbox.pack(side='left', fill='both', expand=True)
    scrollbar.pack(side='right', fill='y')

    # Заполнение списка
    history = load_file_history()
    for i, file_path in enumerate(history, 1):
        listbox.insert(tk.END, f"{i}. {os.path.basename(file_path)}\n{file_path}")

    # Фрейм для кнопок
    button_frame = Frame(main_frame, bg='white')
    button_frame.pack(fill='x', pady=(15, 0))

    # Кнопки с сохранением стиля
    def on_select():
        if listbox.curselection():
            selected_file = history[listbox.curselection()[0]]
            root.selected_file = selected_file
            root.quit()

    select_btn = Button(button_frame,
                        text="Выбрать",
                        command=on_select)
    select_btn.pack(side='left', padx=5)

    def browse_files():
        initial_dir = os.path.dirname(history[0]) if history else ""
        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
            title="Выберите файл Excel"
        )
        if file_path:
            save_file_history(file_path)
            root.selected_file = file_path
            root.quit()

    browse_btn = Button(button_frame,
                        text="Обзор...",
                        command=browse_files)
    browse_btn.pack(side='left', padx=5)

    cancel_btn = Button(button_frame,
                        text="Отмена",
                        command=root.quit)
    cancel_btn.pack(side='right', padx=5)

    root.mainloop()
    root.destroy()
    return getattr(root, 'selected_file', None)


def choose_file():
    """Функция выбора файла с сохранением стиля"""
    history = load_file_history()

    if not history:
        root = tk.Tk()
        apply_style(root)
        root.withdraw()

        file_path = filedialog.askopenfilename(
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
            title="Выберите файл Excel"
        )
        if file_path:
            save_file_history(file_path)
        return file_path
    else:
        return create_history_window()

def read_excel(file_path):
    df = pd.read_excel(file_path)
    data_array = df.values.tolist()
    columns = df.columns.tolist()
    return columns, data_array

def get_webhook_url():
    root = tk.Tk()
    root.withdraw()
    url = simpledialog.askstring("Вебхук URL", "Введите URL для отправки данных через вебхук:")
    return url

def clean_data(data_array):
    cleaned_data = []
    for row in data_array:
        cleaned_row = [
            str(value)[:-2] if pd.notna(value) and str(value).endswith('.0') else str(value) if pd.notna(value) else ''
            for value in row
        ]
        cleaned_data.append(cleaned_row)
    return cleaned_data

def send_data_via_webhook(url, columns, data, progress_bar, progress_label, progress_window):
    failed_payloads = []
    total = len(data)
    start_time = time.time()

    for i, row in enumerate(data):
        payload = {col: row[j] for j, col in enumerate(columns) if row[j] != ''}

        try:
            response = requests.post(url, json=payload)
            if response.status_code != 200:
                raise requests.exceptions.RequestException(f"Response code: {response.status_code}")
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            failed_payloads.append(payload)

        # Обновляем прогресс-бар через метод after
        progress_window.after(0, update_progress, progress_bar, progress_label, i + 1, total)

    if failed_payloads:
        retry_failed_payloads(url, failed_payloads)

    end_time = time.time()
    elapsed_time = end_time - start_time

    # Показываем сообщение через метод after
    progress_window.after(0, show_completion_message, elapsed_time, progress_window)

def update_progress(progress_bar, progress_label, current, total):
    progress_bar['value'] = current / total * 100
    progress_label.config(text=f"Прогресс: {current}/{total} записей отправлено")

def show_completion_message(elapsed_time, progress_window):
    messagebox.showinfo("Информация", f"Отправка данных завершена за {elapsed_time:.2f} секунд.")
    progress_window.quit()  # Завершаем главный цикл Tkinter

def retry_failed_payloads(url, failed_payloads):
    retries = []
    for i, payload in enumerate(failed_payloads):
        try:
            response = requests.post(url, json=payload)
            if response.status_code != 200:
                raise requests.exceptions.RequestException(f"Response code: {response.status_code}")
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            retries.append(payload)

    if retries:
        print("Данные, которые не были успешно отправлены после повторной попытки:")
        for payload in retries:
            print(payload)

def main():
    file_path = choose_file()
    if file_path:
        columns, data = read_excel(file_path)
        cleaned_data = clean_data(data)
        url = get_webhook_url()
        if url:
            progress_window = tk.Tk()
            progress_window.title("Отправка данных")

            progress_label = tk.Label(progress_window, text="Прогресс: 0/0 записей отправлено")
            progress_label.pack(pady=10)

            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=400, mode="determinate")
            progress_bar.pack(pady=20)

            progress_window.geometry("500x200")

            # Запускаем отправку данных в отдельном потоке
            threading.Thread(target=send_data_via_webhook, args=(url, columns, cleaned_data, progress_bar, progress_label, progress_window), daemon=True).start()

            progress_window.mainloop()
        else:
            print("URL не указан")
    else:
        print("Файл не выбран")

if __name__ == "__main__":
    main()