# Extração de Dados de PDF para Excel

Este projeto implementa uma pipeline que extrai dados de uma segunda página de um arquivo PDF e os insere em uma planilha Excel. A função `extrair_dados` é responsável por extrair valores relacionados a palavras ou frases específicas da segunda página de um arquivo PDF e adicionar esses valores à planilha "GERAL" de um arquivo Excel existente.

## Dependências

O projeto requer as seguintes bibliotecas Python:

- `pdfplumber`: Para trabalhar com arquivos PDF.
- `openpyxl`: Para trabalhar com arquivos Excel.

Você pode instalar as dependências usando o comando:

```bash
pip install pdfplumber openpyxl

## Como usar

1. **Prepare o arquivo PDF**: Certifique-se de que o arquivo PDF que você deseja processar está no caminho especificado por `file_path` no código.

2. **Prepare a planilha Excel**: Verifique se a planilha chamada "GERAL" existe no arquivo Excel especificado por `excel_file_path` no código. O código procura pelos cabeçalhos das palavras ou frases específicas na primeira linha da planilha.

3. **Configuração do código**: O código possui uma lista de palavras ou frases específicas em `palavras_frases_especificas` que serão usadas para procurar dados na segunda página do PDF. Altere essa lista conforme necessário para atender às suas necessidades.

4. **Execute o script Python**: Execute o script Python para iniciar a extração de dados do arquivo PDF para a planilha Excel.

5. **Verifique a saída**: Após a execução bem-sucedida do código, os dados extraídos serão adicionados à planilha "GERAL" do arquivo Excel, na próxima linha vazia. A célula correspondente à coluna "Unidade" será preenchida com o valor "Brumado".

## Exemplo de Uso

No código de exemplo fornecido, as seguintes variáveis devem ser configuradas conforme necessário:

- `file_path`: O caminho para o arquivo PDF que você deseja processar.
- `palavras_frases_especificas`: Uma lista de palavras ou frases específicas que você deseja procurar na segunda página do PDF.
- `excel_file_path`: O caminho para o arquivo Excel que você deseja atualizar.

Chame a função `extrair_dados` com essas variáveis para extrair dados do PDF e inseri-los na planilha Excel.

