# Script de Extração de Dados de PDFs para Excel

## Visão Geral

Este script extrai dados de arquivos PDF e insere os resultados em uma planilha Excel. Os dados extraídos incluem informações de interesse, como números relacionados a domínios, aplicações, pacotes, e outras estatísticas relacionadas à segurança de redes.

## Dependências

O script requer as seguintes bibliotecas para funcionar:

- **pdfplumber**: Biblioteca para manipular e extrair texto de arquivos PDF.
- **openpyxl**: Biblioteca para manipular arquivos Excel (XLSX).
- **datetime**: Biblioteca padrão do Python para manipular datas e horários.

As bibliotecas `pdfplumber` e `openpyxl` podem ser instaladas usando o comando pip:

```bash
pip install pdfplumber openpyxl
```

## Funções

### `extrair_dados(pdf_file, buscar, excel_file_path, nome_unidade)`

Função responsável por extrair dados de um arquivo PDF e adicionar esses dados a uma planilha Excel.

- **Parâmetros**:
    - `pdf_file` (`str`): Caminho para o arquivo PDF a ser processado.
    - `buscar` (`list` de `str`): Lista de palavras ou frases que serão buscadas no arquivo PDF.
    - `excel_file_path` (`str`): Caminho para o arquivo Excel onde os dados extraídos serão armazenados.
    - `nome_unidade` (`str`): Nome da unidade correspondente ao arquivo PDF.
- **Descrição**:
    - O script abre o arquivo PDF especificado por `pdf_file`, lê o conteúdo da segunda página, e procura pelas palavras ou frases na lista `buscar`.
    - Se uma palavra ou frase for encontrada em uma linha, a função tenta extrair um número próximo a essa palavra ou frase, e possivelmente uma grandeza associada (GB, TB, MB, KB).
    - Os dados extraídos são armazenados em um dicionário com as palavras ou frases como chaves e os valores extraídos como valores.
    - O script carrega a planilha Excel em `excel_file_path` e adiciona os dados extraídos na próxima linha disponível.
    - Adiciona o nome da unidade e a data atual nas respectivas colunas.
    - Salva o arquivo Excel com os dados atualizados.

## Uso

1. **Configuração**:
    - Defina os caminhos corretos para os arquivos PDF e Excel nos arrays `pdf_paths` e `excel_file_path`, respectivamente.
    - Forneça a lista de palavras ou frases para buscar nos PDFs na variável `buscar`.

2. **Execução**:
    - O script processa cada arquivo PDF na lista `pdf_paths`, extraindo dados de interesse com base nas palavras ou frases fornecidas na lista `buscar`.
    - Os dados extraídos são inseridos na planilha Excel especificada em `excel_file_path` com base no nome da unidade em `unidades`.
    
3. **Resultado**:
    - Após a execução bem-sucedida, o script imprime uma mensagem de confirmação indicando que os dados foram extraídos com sucesso.
    - Os dados extraídos são salvos no arquivo Excel.

## Exemplo

Para executar o script, use o comando a seguir no terminal ou prompt de comando:

```bash
python extrair_dados.py
```
