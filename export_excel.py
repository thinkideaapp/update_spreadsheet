from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def convert_date(date_str):
    month_mapping = {
        'Janeiro': 'January',
        'Fevereiro': 'February',
        'Março': 'March',
        'Abril': 'April',
        'Maio': 'May',
        'Junho': 'June',
        'Julho': 'July',
        'Agosto': 'August',
        'Setembro': 'September',
        'Outubro': 'October',
        'Novembro': 'November',
        'Dezembro': 'December'
    }
    # Converte o nome do mês para inglês usando o dicionário de mapeamento
    for pt_month, en_month in month_mapping.items():
        date_str = date_str.replace(pt_month, en_month)
    # Converte a string da data para um objeto datetime
    date_obj = datetime.strptime(date_str, '%B/%Y')
    # Formata a data para o formato desejado 'YYYY-MM'
    return date_obj.strftime('%Y-%m')


def ler_ultima_linha(planilha):
    # Carrega a planilha
    wb = load_workbook(planilha)
    # Seleciona a primeira planilha
    ws = wb.active
    # Obtém o número total de linhas preenchidas
    num_linhas = ws.max_row
    # Obtém os valores da última linha
    valores_ultima_linha = [ws.cell(
        row=num_linhas, column=col).value for col in range(1, ws.max_column + 1)]
    # Cria um dicionário com base nos valores da última linha
    ultima_linha_dict = {chr(65 + i): valor for i,
                         valor in enumerate(valores_ultima_linha)}
    return ultima_linha_dict


def find_last_row_value(ws, sheet1_dict):
    uc1 = str(sheet1_dict['C']).split('- ')[-1]
    date1 = sheet1_dict['F']
    # Inicia a busca na linha da última linha preenchida
    num_rows = ws.max_row
    while num_rows > 0:
        if ws.cell(row=num_rows, column=2).value == int(uc1):
            date2 = convert_date(str(ws.cell(row=num_rows, column=13).value))
            if date2 == date1:
                return num_rows
            return num_rows
        num_rows -= 1
    # Retorna None se não encontrar o valor
    return None


def get_xlsx_uc(planilha1, planilha2, bill_dict):
    # Lê a última linha da primeira planilha
    last_row_dict = ler_ultima_linha(planilha1)
    uc1 = str(last_row_dict['C']).split('- ')[-1]
    if uc1 != bill_dict['uc']:
        print("UC não encontrada na planilha.")
        return
    # Carrega a segunda planilha
    wb = load_workbook(planilha2)
    # Seleciona a primeira planilha
    ws = wb.active
    # Procura a linha com o valor desejado na coluna B
    last_row = find_last_row_value(ws, last_row_dict)
    if last_row:
        # Se encontrar, insere os valores de G, H e I
        ws.cell(row=last_row, column=column_index_from_string(
            "AD")).value = last_row_dict['G']
        ws.cell(row=last_row, column=column_index_from_string(
            "AE")).value = last_row_dict['H']
        ws.cell(row=last_row, column=column_index_from_string(
            "AF")).value = last_row_dict['I']
        # Salva as alterações na planilha
        wb.save(planilha2)
        print(f"Valores inseridos na linha {last_row}.")
    else:
        print("Valor não encontrado na segunda planilha.")
