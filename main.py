import pandas as pd
import requests
import tkinter as tk
from tkinter import simpledialog, filedialog, ttk, messagebox
import time
import logging
from tqdm import tqdm

# Настройка логирования
logging.basicConfig(filename='webhook_sender.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def choose_file():
    root = tk.Tk()
    root.withdraw()  # Скрыть главное окно

    root.option_add('*Font', 'Arial 14')
    root.option_add('*Background', 'white')
    root.option_add('*Foreground', 'black')
    root.option_add('*Button.Background', 'blue')
    root.option_add('*Button.Foreground', 'white')

    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    logging.info(f'Выбран файл: {file_path}')
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
    logging.info(f'Введён URL: {url}')
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

    for i, row in enumerate(tqdm(data, desc="Отправка данных", unit="запрос", mininterval=1)):
        # Формируем payload, исключая пустые значения
        payload = {col: row[j] for j, col in enumerate(columns) if row[j] != ''}

        try:
            response = requests.post(url, json=payload)
            log_request_response(i + 1, payload, response)
            if response.status_code != 200:
                raise requests.exceptions.RequestException(f"Response code: {response.status_code}")
            response.raise_for_status()  # Проверяем успешность запроса
        except requests.exceptions.RequestException as e:
            logging.error(f"Ошибка при отправке данных: {e}, Данные: {payload}")
            failed_payloads.append(payload)
        time.sleep(0.2)  # Ожидание 0.2 секунды

        # Обновляем прогресс-бар
        progress_bar['value'] = (i + 1) / total * 100
        progress_label.config(text=f"Прогресс: {i + 1}/{total} записей отправлено")
        progress_bar.update()

    if failed_payloads:
        retry_failed_payloads(url, failed_payloads)

    end_time = time.time()
    elapsed_time = end_time - start_time
    messagebox.showinfo("Информация", f"Отправка данных завершена за {elapsed_time:.2f} секунд.")

    progress_window.destroy()


def retry_failed_payloads(url, failed_payloads):
    retries = []
    for i, payload in enumerate(failed_payloads):
        try:
            response = requests.post(url, json=payload)
            log_request_response(i + 1, payload, response)
            if response.status_code != 200:
                raise requests.exceptions.RequestException(f"Response code: {response.status_code}")
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            logging.error(f"Повторная ошибка при отправке данных: {e}, Данные: {payload}")
            retries.append(payload)
        time.sleep(2)  # Ожидание 2 секунды между повторными попытками

    if retries:
        logging.error("Данные, которые не были успешно отправлены после повторной попытки:")
        for payload in retries:
            logging.error(payload)


def log_request_response(index, request, response):
    logging.info(f"Запрос {index}: {request}")
    logging.info(f"Ответ {index}: {response.status_code} {response.text}")


def main():
    file_path = choose_file()
    if file_path:
        columns, data = read_excel(file_path)
        cleaned_data = clean_data(data)
        url = get_webhook_url()
        if url:
            # Создаем диалоговое окно с прогресс-баром
            progress_window = tk.Tk()
            progress_window.title("Отправка данных")

            progress_label = tk.Label(progress_window, text="Прогресс: 0/0 записей отправлено")
            progress_label.pack(pady=10)

            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=400, mode="determinate")
            progress_bar.pack(pady=20)

            progress_window.geometry("500x200")

            # Запускаем отправку данных и обновление прогресс-бара
            progress_window.after(100, send_data_via_webhook, url, columns, cleaned_data, progress_bar, progress_label,
                                  progress_window)

            progress_window.mainloop()
        else:
            logging.warning("URL не указан")
    else:
        logging.warning("Файл не выбран")


if __name__ == "__main__":
    main()