import pandas as pd
import requests
import tkinter as tk
from tkinter import simpledialog, filedialog, ttk, messagebox
import time
import threading
import json
import os

# Путь к файлу для сохранения последнего выбранного пути
CONFIG_FILE = "last_path.json"

def load_last_path():
    """Загружает последний выбранный путь из файла."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as file:
            return json.load(file).get("last_path", "")
    return ""

def save_last_path(path):
    """Сохраняет последний выбранный путь в файл."""
    with open(CONFIG_FILE, "w") as file:
        json.dump({"last_path": os.path.dirname(path)}, file)

def choose_file():
    root = tk.Tk()
    root.withdraw()  # Скрыть главное окно

    root.option_add('*Font', 'Arial 14')
    root.option_add('*Background', 'white')
    root.option_add('*Foreground', 'black')
    root.option_add('*Button.Background', 'blue')
    root.option_add('*Button.Foreground', 'white')

    # Загружаем последний выбранный путь
    initial_dir = load_last_path()

    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
        initialdir=initial_dir  # Указываем последний выбранный путь
    )

    if file_path:
        save_last_path(file_path)  # Сохраняем новый путь

    return file_path

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