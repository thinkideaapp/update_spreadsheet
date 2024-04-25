import datetime
import re

import openpyxl
from openpyxl.styles import Alignment, Font
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


def search_text_sum(text):
    sums_dict = {
        "sum_ac": True,
        "sum_ab": True,
        "sum_z": True,
        "sum_z_b": True,
    }

    text_find1 = text.find("ENERGIA ATIVA FORNECIDA FP")
    text_find2 = text.find("ENERGIA ATIVA FORNECIDA FP -")
    if text_find1 == -1 and text_find2 == -1:
        sums_dict["sum_ac"] = False
    text_find1 = text.find("ENERGIA ATIVA FORNECIDA P")
    text_find2 = text.find("ENERGIA ATIVA FORNECIDA P -")
    if text_find1 == -1 and text_find2 == -1:
        sums_dict["sum_ab"] = False
    text_find1 = text.find("DEMANDA ULTRAPASSAGEM")
    text_find2 = text.find("INDEN. VIOL. PRAZO ATENDIMENTO")
    if text_find1 == -1 and text_find2 == -1:
        sums_dict["sum_z"] = False
    text_find1 = text.find("JUROS MORATÓRIA")
    text_find2 = text.find("MULTA -")
    if text_find1 == -1 and text_find2 == -1:
        sums_dict["sum_z_b"] = False
    return sums_dict


def get_info_rows(text, get_type):
    text_init = "ENERGIA ATIVA FORNECIDA"
    text_end = "DEMANDA - kW"
    if get_type == "kwh_consumed":
        text_init = "ENERGIA GERAÇÃO"
    index_init = text.find(text_init)
    index_end = text.find(text_end)
    text_part = text[index_init:index_end]
    rows = text_part.split('\n')

    quantity_rows = [
        "ENERGIA ATIVA FORNECIDA FP",
        "ENERGIA ATIVA FORNECIDA HR",
        "ENERGIA ATIVA FORNECIDA P",
        "ENERGIA INJETADA FP",
        "ENERGIA INJETADA HR",
        "ENERGIA INJETADA P",
    ]
    unit_prices_rows = [
        "ENERGIA ATIVA FORNECIDA FP",
        "ENERGIA ATIVA FORNECIDA P",
        "ENERGIA ATIVA FORNECIDA FP -",
        "ENERGIA ATIVA FORNECIDA P -",
    ]
    prices_rows = [
        "DEMANDA",
        "DEMANDA ULTRAPASSAGEM",
        "UFER FP",
        "CONTRIB. ILUM. PÚBLICA",
        "INDEN. VIOL. PRAZO ATENDIMENTO",
    ]
    kwh_rows = rows[:3]

    row_list = []
    value = None
    for i, row in enumerate(rows, start=1):
        if get_type == "quantity":
            if row.split(' kWh')[0] in quantity_rows:
                if len(row.split(' kWh')[0].split(" ")) == 4:
                    value = row.split(' ')[6]
                else:
                    value = row.split(' ')[5]
                if value:
                    row_list.append(value)
                    quantity_rows.remove(row.split(' kWh')[0])
            elif row.split(' -')[0] in quantity_rows:
                if len(row.split(' -')[0].split(" ")) == 4:
                    value = row.split(' ')[8]
                else:
                    value = row.split(' ')[7]
                row_list.append(value)
                quantity_rows.remove(row.split(' -')[0])
        elif get_type == "unit_price":
            if row.split(' kWh')[0] in unit_prices_rows:
                if len(row.split(' kWh')[0].split(" ")) == 4:
                    value = row.split(' ')[5]
                else:
                    value = row.split(' ')[8]
                row_list.append(value)
            if row.split(' -')[0] in unit_prices_rows:
                if len(row.split(' kWh')[0].split(" ")) == 6:
                    value = row.split(' ')[7]
                else:
                    value = row.split(' ')[8]
                row_list.append(value)
        elif get_type == "prices":
            if row.split(' kW')[0] in prices_rows:
                if len(row.split(' kW')[0].split(" ")) == 1:
                    value = row.split(' ')[5]
                else:
                    value = row.split(' ')[6]
                row_list.append(value)
                prices_rows.remove(row.split(' kW')[0])
                continue
            if row.split(' kVArh')[0] == "UFER FP":
                value = row.split(' ')[-1]
                row_list.append(value)
                continue
            if row.split(' -')[0] == prices_rows[-1]:
                value = row.split(' ')[-1]
                row_list.append(value)
                prices_rows.remove(row.split(' -')[0])
            if row.split(' -')[0] in prices_rows:
                value = re.sub("[a-zA-Z]", "", row.split(" ")[-2])
                row_list.append(value)
                prices_rows.remove(row.split(' -')[0])
        elif get_type == "kwh_consumed":
            if row in kwh_rows:
                value = row.split(' ')[5]
                row_list.append(value)
    return row_list


def bill_classification(text):
    # Expressão regular para encontrar a letra após "Classificação: "
    text_default = r'Classificação:\s*([A-Z])'

    # Procurar o padrão na linha
    match = re.search(text_default, text)

    if match:
        # Retornar a letra encontrada
        return match.group(1)
    else:
        # Caso não encontre o padrão, retornar None ou uma mensagem
        return None


def sum_values(unit_prices, prices, bill_group, sums_dict):
    if bill_group == "A":
        if sums_dict["sum_ac"]:
            unit_prices[0] = unit_prices[0]+unit_prices[2]
        if sums_dict["sum_ab"]:
            unit_prices[1] = unit_prices[1]+unit_prices[3]
        if sums_dict["sum_z"]:
            prices[1] = prices[1]+prices[-1]
    else:
        unit_prices = unit_prices
        if sums_dict["sum_z_b"]:
            prices[1] = prices[1]+prices[-1]
    return unit_prices, prices


def find_values(text):
    bill_group = bill_classification(text)
    # Definindo padrões de expressões regulares para o valor e a data
    quantity = get_info_rows(text, "quantity")
    unit_prices = get_info_rows(text, "unit_price")
    prices = get_info_rows(text, "prices")
    kwh_consumed = get_info_rows(text, "kwh_consumed")

    price_default = r'R\$.*?(\d{1,3}(?:\.\d{3})*,\d{2})'
    date_default = r'(\d{2}/\d{2}/\d{4})'
    interval_default = r'(\d{2}/\d{2}/\d{4})'
    uc_default = r'UC (\d+)'

    # Procurando pelo padrão do valor
    match_valor = re.search(price_default, text)
    match_date = re.findall(date_default, text)
    match_uc = re.search(uc_default, text)
    match_interval = re.findall(interval_default, text)

    if match_valor:
        price = match_valor.group(1)
    if match_date:
        date = match_date[-1]
        date = format_date(date)
    if match_interval:
        cycle = f"{match_interval[2]} a {match_interval[3]}"
    if match_uc:
        uc = match_uc.group(1)

    unit_prices = [number.replace('.', '').replace(',', '.')
                   for number in unit_prices]
    unit_prices = [float(number) for number in unit_prices]

    prices = [number.replace('.', '').replace(',', '.') for number in prices]
    prices = [float(number) for number in prices]

    quantity = [number.replace('.', '').replace(',', '.')
                for number in quantity]
    quantity = [float(number) for number in quantity]
    kwh_consumed = [number.replace('.', '').replace(',', '.')
                    for number in kwh_consumed]
    kwh_consumed = [float(number) for number in kwh_consumed]
    price = str(price).replace('.', '').replace(',', '.')

    sums_dict = search_text_sum(text)
    unit_prices, prices = sum_values(unit_prices, prices, bill_group, sums_dict)

    bill_dict = {
        'bill_group': bill_group,
        'price': price,
        'date': date,
        'cycle': cycle,
        'uc': uc,
        'quantity': quantity,
        'unit_price': unit_prices,
        'prices': prices,
        'kwh_consumed': kwh_consumed
    }
    return bill_dict


def organize_sheet_columns(sheet, max_row, bill_dict):
    font_trebuchet_ms = Font(name='Trebuchet MS')

    center_cell = sheet.cell(row=max_row, column=column_index_from_string(
        'M'), value=bill_dict['date'])
    center_cell.alignment = Alignment(horizontal='center')
    center_cell.font = font_trebuchet_ms

    center_cell = sheet.cell(row=max_row, column=column_index_from_string(
        'N'), value=bill_dict['cycle'])
    center_cell.alignment = Alignment(horizontal='center')
    center_cell.font = font_trebuchet_ms

    price_cell = sheet.cell(row=max_row, column=column_index_from_string(
        'O'), value=float(bill_dict['price']))
    price_cell.number_format = 'R$ #,##0.00'
    price_cell.font = font_trebuchet_ms

    center_cell = sheet.cell(
        row=max_row, column=column_index_from_string('B'), value=int(
            bill_dict['uc']))
    center_cell.alignment = Alignment(horizontal='center')
    center_cell.font = font_trebuchet_ms

    if bill_dict['bill_group'] == 'A':
        if len(bill_dict['quantity']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AO'), value=float(bill_dict['quantity'][0]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AP'), value=float(bill_dict['quantity'][1]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 3:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AN'), value=float(bill_dict['quantity'][2]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 4:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AX'), value=float(bill_dict['quantity'][3]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 5:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AY'), value=float(bill_dict['quantity'][4]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 6:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AW'), value=float(bill_dict['quantity'][5]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['unit_price']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AC'), value=float(round(bill_dict['unit_price'][0], 5)))
            price_cell.number_format = 'R$ #,##0.00000'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['unit_price']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AB'), value=float(bill_dict['unit_price'][1]))
            price_cell.number_format = 'R$ #,##0.00000'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'Q'), value=float(bill_dict['prices'][0]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'Z'), value=float(bill_dict['prices'][1]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 3:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'S'), value=float(bill_dict['prices'][2]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 4:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'U'), value=float(bill_dict['prices'][3]))
            price_cell.number_format = 'R$ #,##0.00'
        price_cell.font = font_trebuchet_ms

        if len(bill_dict['kwh_consumed']) >= 1:
            sheet.cell(row=max_row, column=column_index_from_string(
                'AT'), value=float(bill_dict['kwh_consumed'][0]))
        if len(bill_dict['kwh_consumed']) >= 2:
            sheet.cell(row=max_row, column=column_index_from_string(
                'AS'), value=float(bill_dict['kwh_consumed'][1]))
        if len(bill_dict['kwh_consumed']) >= 3:
            sheet.cell(row=max_row, column=column_index_from_string(
                'AU'), value=float(bill_dict['kwh_consumed'][2]))
    else:
        if len(bill_dict['quantity']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AO'), value=float(bill_dict['quantity'][0]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['quantity']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AX'), value=float(bill_dict['quantity'][1]))
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['unit_price']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AC'), value=float(bill_dict['unit_price'][0]))
            price_cell.number_format = 'R$ #,##0.00000'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['unit_price']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'AB'), value=float(bill_dict['unit_price'][0]))
            price_cell.number_format = 'R$ #,##0.00000'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 1:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'Z'), value=float(bill_dict['prices'][1]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['prices']) >= 2:
            price_cell = sheet.cell(row=max_row, column=column_index_from_string(
                'U'), value=float(bill_dict['prices'][0]))
            price_cell.number_format = 'R$ #,##0.00'
            price_cell.font = font_trebuchet_ms

        if len(bill_dict['kwh_consumed']) >= 1:
            sheet.cell(row=max_row, column=column_index_from_string(
                'AT'), value=float(bill_dict['kwh_consumed'][0]))


def insert_sheet(sheet_path, bill_dict):
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
    organize_sheet_columns(sheet, max_row, bill_dict)

    workbook.save(sheet_path)
    print(f"Valor inserido na planilha {sheet_path}")
