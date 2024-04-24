import os
import time

import PyPDF2
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from utils import find_values, format_date, insert_sheet, pdf_text


class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        # Verifica se o arquivo é um PDF
        if event.src_path.endswith('.pdf'):
            print(f'Arquivo PDF detectado: {event.src_path}')
            bill_dict = self.read_pdf(event.src_path)
            sheet_path = "planilha.xlsx"
            insert_sheet(sheet_path, bill_dict)
            os.remove(event.src_path)
            print(f'Arquivo {event.src_path} removido')
            print('Monitorando a pasta...')

    def read_pdf(self, file_path):
        time.sleep(1)
        try:
            with open(file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                print(f'Número de páginas: {len(reader.pages)}')

                for page in reader.pages:
                    text = page.extract_text()
                    # text = pdf_text(text)
                    bill_dict = find_values(text)
                    if bill_dict:
                        break

                return bill_dict
        except Exception as e:
            print(f"Não foi possível ler o arquivo {file_path}: {e}")


def main():
    print("Monitorando a pasta...")
    path = '.'  # Define o caminho da pasta a ser monitorada
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()