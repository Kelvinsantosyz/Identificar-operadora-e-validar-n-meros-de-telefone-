from openpyxl import load_workbook
from phonenumbers import carrier,parse,is_valid_number,phonenumberutil,NumberParseException
# Abre a planilha do Excel
workbook = load_workbook('planilha.xlsx')

# Seleciona a planilha ativa
worksheet = workbook.active

# Percorre as células da coluna A
phone_number = None
for cell in worksheet['A']:
    phone_number = cell.value
    # Verifica se o valor é um número de telefone válido
    if phone_number is None:
        print("A célula está vazia")
        continue
    try:
        phone_number = parse(phone_number)
        if not is_valid_number(phone_number):
            print(f'O número {phone_number} não é válido')
            continue
    except phonenumberutil.NumberParseException:
        print(f'O valor {phone_number} não é um número de telefone válido')
        continue
    except Exception:
        print("Ocorreu erro inesperado ")

    # Obtém o nome da operadora correspondente ao número de telefone
    carrier_name = carrier.name_for_number(phone_number, 'pt-BR')
    print(f'O número {phone_number} é válido e a operadora é {carrier_name}')
