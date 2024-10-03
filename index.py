import openpyxl

# Carregar o arquivo Excel
caminho_arquivo = r'C:\Users\asoares\Documents\testePY.xlsx'
workbook = None  # Inicializando workbook como None

try:
    workbook = openpyxl.load_workbook(caminho_arquivo)
    sheet = workbook.active

    # Definir o intervalo de linhas a ser processado
    inicio_linha = 2
    fim_linha = 1270

    # Coletar valores distintos da coluna O
    valores_distintos = set()
    for linha in range(inicio_linha, fim_linha + 1):
        valor = sheet[f'O{linha}'].value
        if valor is not None:
            valores_distintos.add(valor)

    # Converter o conjunto em lista
    lista_valores_distintos = list(valores_distintos)
    print("Valores distintos da coluna O:", lista_valores_distintos)

    # Loop para iterar pelas linhas do Excel
    for linha in range(inicio_linha, fim_linha + 1):
        selecao_atual = sheet[f'O{linha}'].value
        placa_atual = sheet[f'H{linha}'].value
        
        # Garantir que placa_atual é um número
        try:
            placa_atual = int(placa_atual)
        except (ValueError, TypeError):
            continue  # Ignorar se não for um número

        # Verificar as condições para 'SelecaoAtual' e 'PlacaAtual'
        if placa_atual <= 16:
            if selecao_atual == '22120001_CONV':
                sheet[f'F{linha}'] = 1
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120001_SEG':
                sheet[f'F{linha}'] = 2
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120002_CONV':
                sheet[f'F{linha}'] = 3
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120002_SEG':
                sheet[f'F{linha}'] = 4
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120014_CONV':
                sheet[f'F{linha}'] = 5
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120014_IPRO':
                sheet[f'F{linha}'] = 6
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120017_CONV':
                sheet[f'F{linha}'] = 7
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120017_IPRO':
                sheet[f'F{linha}'] = 8
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120023_CONV':
                sheet[f'F{linha}'] = 9
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120023_IPRO':
                sheet[f'F{linha}'] = 10
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120070_CONV':
                sheet[f'F{linha}'] = 11
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120070_IPRO':
                sheet[f'F{linha}'] = 12
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120116_CONV':
                sheet[f'F{linha}'] = 13
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120121_CONV':
                sheet[f'F{linha}'] = 14
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120121_IPRO':
                sheet[f'F{linha}'] = 15
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120125_CONV':
                sheet[f'F{linha}'] = 16
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120125_SEG':
                sheet[f'F{linha}'] = 17
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120130_CONV':
                sheet[f'F{linha}'] = 18
                sheet[f'P{linha}'] = 'RODADA 1'
            elif selecao_atual == '22120131_CONV':
                sheet[f'F{linha}'] = 19
                sheet[f'P{linha}'] = 'RODADA 1'

    # Salvar as alterações no arquivo Excel
    workbook.save(caminho_arquivo)
    print("Automação concluída com sucesso!")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

finally:
    if workbook is not None:
        workbook.close()