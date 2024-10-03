import openpyxl

caminho_arquivo = r'C:\Users\asoares\Documents\CHIPPING - B_CK_23_20_BC1F3 - M4 - 24410663.xlsx'
workbook = None

try:
    workbook = openpyxl.load_workbook(caminho_arquivo)
    sheet = workbook.active

    # Definir o intervalo de linhas a ser processado
    inicio_linha = 2
    fim_linha = 3363

    # Dicionário para armazenar listas de valores distintos da coluna O por intervalos
    intervalos_valores = {}

    # Inicializa o intervalo de placas
    intervalo_atual = 16
    rodada_global = 1  # Iniciar a contagem da rodada global

    while True:
        # Cria uma chave para o intervalo atual
        intervalo_chaves = f'{intervalo_atual - 15} - {intervalo_atual}'
        intervalos_valores[intervalo_chaves] = []

        for linha in range(inicio_linha, fim_linha + 1):
            placa_atual = sheet[f'H{linha}'].value
            
            # Parar se a célula estiver em branco
            if placa_atual is None:
                break
            
            try:
                placa_atual = int(placa_atual)
            except (ValueError, TypeError):
                continue  # Ignorar se não for um número
            
            # Verifica se a placa atual está dentro do intervalo atual
            if intervalo_atual - 15 <= placa_atual <= intervalo_atual:
                # Seleção anterior era 'O'
                selecao_atual = sheet[f'N{linha}'].value
                if selecao_atual is not None and selecao_atual not in intervalos_valores[intervalo_chaves]:
                    intervalos_valores[intervalo_chaves].append(selecao_atual)
            elif placa_atual > intervalo_atual:
                # Se a placa atual for maior que o intervalo atual, verifica o próximo intervalo
                break

        # Verifica se existem valores para o intervalo atual
        if not intervalos_valores[intervalo_chaves]:
            del intervalos_valores[intervalo_chaves]  # Remove o intervalo se não tiver valores
            break

        # Atualiza o intervalo para o próximo
        intervalo_atual += 16

    # Exibir resultados
    print("Valores distintos da coluna O por intervalos de placas:")
    for intervalo, valores in intervalos_valores.items():
        print(f"{intervalo}: {valores}")

    # Atribuir seleções a rodadas e salvar no dicionário
    rodadas = {}
    rodada_global = 1  # Contador global de rodadas

    for intervalo, selecoes in intervalos_valores.items():
        num_selecoes = len(selecoes)
        if num_selecoes == 0:
            continue

        # Para cada intervalo, inicia na rodada atual (não reinicia em 1)
        rodada_atual = rodada_global
        selecao_index = 1

        for i in range(num_selecoes):
            # Se atingirmos 5 seleções, incrementa a rodada
            if selecao_index > 5:
                rodada_atual += 1
                selecao_index = 1
                rodada_global += 1  # Atualiza a rodada global

            # Inicializa o dicionário rodadas se não existir
            if intervalo not in rodadas:
                rodadas[intervalo] = {}

            # Inicializa a rodada para o intervalo se não existir
            if rodada_atual not in rodadas[intervalo]:
                rodadas[intervalo][rodada_atual] = []

            # Adiciona a seleção à rodada correspondente
            rodadas[intervalo][rodada_atual].append({
                "seleção": selecoes[i],
                "seleção_numero": selecao_index
            })
            
            selecao_index += 1

        # Atualiza a rodada global após processar cada intervalo
        rodada_global = rodada_atual + 1

    # Exibir resultados
    print("Seleções atribuídas a rodadas por intervalos de placas:")
    for intervalo, rodadas_lista in rodadas.items():
        for rodada, selecoes in rodadas_lista.items():
            for selecao_info in selecoes:
                print(f"{intervalo}, Rodada {rodada}, seleção {selecao_info['seleção_numero']}: {selecao_info['seleção']}")

    # Iterar em cada linha da planilha para preencher as colunas F e P
    for linha in range(inicio_linha, fim_linha + 1):
        placa_atual = sheet[f'H{linha}'].value
        
        # Parar se a célula estiver em branco
        if placa_atual is None:
            break

        try:
            placa_atual = int(placa_atual)
        except (ValueError, TypeError):
            continue  # Ignorar se não for um número
        
        # Checa em qual intervalo a placa atual está
        intervalo_encontrado = None
        for intervalo in intervalos_valores.keys():
            inicio_intervalo, fim_intervalo = map(int, intervalo.split(' - '))
            if inicio_intervalo <= placa_atual <= fim_intervalo:
                intervalo_encontrado = intervalo
                break
        
        # Se o intervalo foi encontrado, preenche as colunas F e P
        if intervalo_encontrado:
            selecao_atual = sheet[f'N{linha}'].value
            if selecao_atual:
                for rodada, selecoes in rodadas[intervalo_encontrado].items():
                    for selecao_info in selecoes:
                        if selecao_info["seleção"] == selecao_atual:
                            # Preenche a coluna F com o número da seleção (1, 2, 3, 4 ou 5)
                            sheet[f'F{linha}'] = selecao_info["seleção_numero"]
                            # Preenche a coluna P com o valor da rodada
                            sheet[f'O{linha}'] = f'Rodada {rodada}'
                            break

    # Salvar as alterações no arquivo Excel
    workbook.save(caminho_arquivo)
    print("Automação concluída com sucesso!")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

finally:
    if workbook is not None:
        workbook.close()
