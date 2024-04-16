import datetime
import re

import openpyxl
from openpyxl.utils import column_index_from_string


def format_date(date):
    # Convert date from "DD/MM/YYYY" to "MARÇO/YYYY"
    date_obj = datetime.datetime.strptime(date, '%d/%m/%Y')
    months = [
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio',
        'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
        'Novembro', 'Dezembro'
    ]
    month_str = months[date_obj.month - 1]
    formatted_date = f"{month_str}/{date_obj.year}"
    return formatted_date


def pdf_text(text):
    inicio = text.find("R$")

    if inicio == -1:
        return None

    fim = text.find("\n", inicio)

    if fim != -1:
        return text[inicio:fim]
    else:
        return text[inicio:]


def find_values(text):
    # Definindo padrões de expressões regulares para o valor e a data
    price_default = r'R\$.*?(\d{1,3}(?:\.\d{3})*,\d{2})'
    date_default = r'(\d{2}/\d{2}/\d{4})'
    interval_default = r'(\d{2}/\d{2}/\d{4})'

    # Procurando pelo padrão do valor
    match_valor = re.search(price_default, text)
    if match_valor:
        price = match_valor.group(1)
    else:
        price = None

    # Procurando pelo padrão da data (no final da string)
    match_date = re.findall(date_default, text)
    date = match_date[-1] if match_date else None
    match_interval = re.findall(interval_default, text)
    date_match_interval = match_interval[1:3]
    date_interval = f"{date_match_interval[0]} a {date_match_interval[1]}"
    return price, date, date_interval


def insert_sheet(sheet_path, price, date, date_interval):
    workbook = openpyxl.load_workbook(sheet_path)
    sheet = workbook.active

    value_column = column_index_from_string("O")
    max_row = sheet.max_row
    for linha in range(max_row, 0, -1):
        # Obter o valor da célula na coluna especificada
        valor = sheet.cell(row=linha, column=value_column).value

        # Verificar se a célula contém um valor não vazio
        if valor is not None:
            # Retornar a linha em que encontrou um valor não vazio
            max_row = linha + 1
            break

    date_column = column_index_from_string("M")
    interval_column = column_index_from_string("N")
    price_column = column_index_from_string("O")

    sheet.cell(row=max_row, column=date_column, value=date)
    sheet.cell(row=max_row, column=interval_column, value=date_interval)
    price_cell = sheet.cell(row=max_row, column=price_column, value=price)
    price_cell.number_format = 'R$ #,##0.00'

    workbook.save(sheet_path)
    print(f"Valor inserido na planilha {sheet_path}")
