import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter

# Função para extrair dados da segunda página do PDF
def extrair_dados(pdf_file, palavras_frases_especificas, excel_file_path):
    # Inicializar um dicionário com palavras ou frases específicas como chaves e valores como 0
    dados_extraidos = {palavra_ou_frase: 0 for palavra_ou_frase in palavras_frases_especificas}
    grandezas = ["GB", "TB", "MB", "KB"]  # Adicione aqui todas as grandezas possíveis

    # Abrir o arquivo PDF
    with pdfplumber.open(pdf_file) as pdf:
        # Verificar se há pelo menos duas páginas
        if len(pdf.pages) >= 2:
            # Acessar a segunda página (índice 1)
            segunda_pagina = pdf.pages[1]
            # Extraia o texto da segunda página
            texto = segunda_pagina.extract_text()
            # Divida o texto em linhas
            linhas = texto.split('\n')
            
            # Processar cada linha
            for linha in linhas:
                # Iterar sobre as palavras ou frases específicas
                for palavra_ou_frase in palavras_frases_especificas:
                    # Verificar se a linha contém a palavra ou frase específica
                    if palavra_ou_frase in linha:
                        # Dividir a linha em palavras
                        palavras = linha.split()
                        
                        # Iterar pelas palavras na linha
                        encontrou_numero = False
                        for i, palavra in enumerate(palavras):
                            # Verificar se a palavra é um número (inteiro ou float)
                            try:
                                # Tentar converter a palavra para float
                                numero = float(palavra)
                                # Se a conversão for bem-sucedida, verificar a palavra ou frase anterior
                                if i > 0:
                                    # Verifica o texto anterior para ver se contém a palavra ou frase específica
                                    texto_antes_do_numero = ' '.join(palavras[i - len(palavra_ou_frase.split()):i])
                                    
                                    # Comparar a palavra ou frase específica com o texto anterior
                                    if texto_antes_do_numero == palavra_ou_frase:
                                        # Verifique a próxima palavra após o número para verificar a grandeza
                                        grandeza = ""
                                        if i + 1 < len(palavras):
                                            proxima_palavra = palavras[i + 1]
                                            # Verifique se a próxima palavra é uma grandeza conhecida
                                            if proxima_palavra in grandezas:
                                                grandeza = proxima_palavra
                                        
                                        # Combine o número com a grandeza, separando-os por um espaço
                                        dados_extraidos[palavra_ou_frase] = f"{numero} {grandeza}"
                                        encontrou_numero = True
                                        break
                            except ValueError:
                                # Ignorar palavras que não são números
                                continue
                        
                        # Se não encontrou número para a palavra ou frase específica, o valor permanece 0
    # Carregar o arquivo Excel usando openpyxl
    workbook = openpyxl.load_workbook(excel_file_path)

    # Selecionar a planilha chamada "GERAL"
    if "Geral" in workbook.sheetnames:
        sheet = workbook["Geral"]
    else:
        raise ValueError("A planilha 'GERAL' não existe no arquivo Excel.")

    # Encontrar a próxima linha vazia
    proxima_linha = sheet.max_row + 1

    # Adicionar "Brumado" na coluna "Unidade"
    coluna_unidade = None
    for coluna in range(1, sheet.max_column + 1):
        # Obtenha a letra da coluna para a célula da primeira linha
        coluna_letra = get_column_letter(coluna)
        # Obtenha a célula do cabeçalho na primeira linha
        celula = sheet[f"{coluna_letra}1"]
        
        # Verifique se o cabeçalho da coluna é "Unidade"
        if celula.value == "Unidade":
            # Adicione "Brumado" à próxima linha vazia da coluna "Unidade"
            sheet[f"{coluna_letra}{proxima_linha}"].value = "Brumado"
            break

    # Adicionar os dados extraídos à planilha "GERAL"
    for palavra_ou_frase in palavras_frases_especificas:
        # Encontrar a coluna correspondente à palavra ou frase na planilha "GERAL"
        coluna_encontrada = False
        for coluna in range(1, sheet.max_column + 1):
            # Obtenha a letra da coluna para a célula da primeira linha
            coluna_letra = get_column_letter(coluna)
            # Obtenha a célula do cabeçalho na primeira linha
            celula = sheet[f"{coluna_letra}1"]
            
            # Verifique se o cabeçalho da coluna corresponde à palavra ou frase
            if celula.value == palavra_ou_frase:
                # Adicione o dado extraído à célula correspondente na próxima linha vazia
                nova_celula = sheet[f"{coluna_letra}{proxima_linha}"]   
                nova_celula.value = dados_extraidos[palavra_ou_frase]
                
                coluna_encontrada = True
                break
        
        # Se a palavra ou frase não foi encontrada como cabeçalho de coluna
        if not coluna_encontrada:
            pass
    
    # Salvar o arquivo Excel com os dados atualizados
    workbook.save(excel_file_path)

    return dados_extraidos

# Exemplo de uso
file_path = '../test_pip/test.pdf'

palavras_frases_especificas = ['Web domains accessed', 'Applications accessed', 'Blocked Applications', 'Intrusion Attacks', 'Emergency + Critical Attacks', 'Application Data Transfer', 'Web data transfer', 'Total User Data Transfer', 'App Risk Score (out of 5)']

excel_file_path = 'test.xlsx'  

# Chamar a função para extrair dados
dados_extraidos = extrair_dados(file_path, palavras_frases_especificas, excel_file_path)

print("Dados extraídos com sucesso!")
