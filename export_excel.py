from datetime import datetime
import locale

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def convert_date(date_str, reverse=False):

    if reverse:
        month_mapping = {
            "January": "Janeiro",
            "February": "Fevereiro",
            "March": "Março",
            "April": "Abril",
            "May": "Maio",
            "June": "Junho",
            "July": "Julho",
            "August": "Agosto",
            "September": "Setembro",
            "October": "Outubro",
            "November": "Novembro",
            "December": "Dezembro"
        }
        # Separando o ano e o mês da string da data
        ano, mes_numero = date_str.split('-')

        # Convertendo o número do mês para o nome do mês
        mes_nome = list(month_mapping.keys())[int(mes_numero) - 1]

        # Montando a string da data no formato desejado
        data_formatada = f"{month_mapping[mes_nome]}/{ano}"

        return data_formatada
    else:
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


def read_last_row(planilha, uc, date):
    # Carregar a planilha
    wb = load_workbook(planilha)

    # Selecionar a primeira planilha
    planilha = wb.active

    # Encontrar o número de linhas e colunas na planilha
    num_linhas = planilha.max_row
    num_colunas = planilha.max_column

    # Inicializar um dicionário vazio
    dados_dict = {}

    # Iterar sobre as linhas da planilha
    for linha in range(1, num_linhas + 1):
        valor_b = str(planilha.cell(row=linha, column=3).value).split('- ')[-1]
        valor_f = convert_date(
            str(planilha.cell(row=linha, column=6).value), True)

        # Verificar se o valor na coluna B é igual a 109228
        if valor_b == uc and valor_f == date:
            # Iterar sobre as colunas da última linha
            for coluna in range(1, num_colunas + 1):
                # Adicionar os valores da última linha ao dicionário
                chave = planilha.cell(row=linha, column=coluna).value
                dados_dict[f'{coluna}'] = chave
            print(f'Valor encontrado na linha: {linha}')
    return dados_dict


def find_last_row_value(ws, uc1, date1=None, reverse=False):
    # Inicia a busca na linha da última linha preenchida
    num_rows = ws.max_row
    while num_rows > 0:
        if ws.cell(row=num_rows, column=2).value == int(uc1):
            try:
                date2 = convert_date(
                    str(ws.cell(row=num_rows, column=13).value), reverse)
            except:
                date2 = str(ws.cell(row=num_rows, column=13).value)
            if date1:
                if date2 == date1:
                    return num_rows
            else:
                return num_rows
        num_rows -= 1
    # Retorna None se não encontrar o valor
    return None


def get_xlsx_uc(planilha1, planilha2, bill_dict):
    # Lê a última linha da primeira planilha
    last_row_dict = read_last_row(
        planilha1, bill_dict['uc'], bill_dict['date'])
    # Carrega a segunda planilha
    wb = load_workbook(planilha2)
    # Seleciona a primeira planilha
    ws = wb.active
    # Procura a linha com o valor desejado na coluna B
    last_row = find_last_row_value(
        ws, bill_dict['uc'], bill_dict['date'], True)
    if last_row and last_row_dict:
        # Se encontrar, insere os valores de G, H e I
        ws.cell(row=last_row, column=column_index_from_string(
            "AD")).value = last_row_dict['7']
        ws.cell(row=last_row, column=column_index_from_string(
            "AE")).value = last_row_dict['8']
        ws.cell(row=last_row, column=column_index_from_string(
            "AF")).value = last_row_dict['9']
        # Salva as alterações na planilha
        wb.save(planilha2)
    else:
        print("Valor não encontrado na segunda planilha.")
