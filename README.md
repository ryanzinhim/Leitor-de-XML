# Processador de XMLs para Excel

Este script Python foi desenvolvido para processar **grandes quantidades de arquivos XML** de Notas Fiscais Eletrônicas (NF-e) e extrair informações relevantes, como o número da NF e a chave de acesso. Essas informações são então salvas em uma planilha do Excel. O script é otimizado para lidar com **volumes massivos de dados** de forma eficiente.

---

## Funcionalidades

- **Leitura de XMLs**: O script lê todos os arquivos XML em um diretório específico, suportando **milhares de arquivos**.
- **Extração de Dados**: Extrai o número da NF e a chave de acesso de cada XML.
- **Verificação de Duplicatas**: Identifica e remove números de NF duplicados, garantindo a integridade dos dados.
- **Verificação de Faltantes**: Verifica se há números de NF faltantes em um intervalo especificado.
- **Salvamento em Excel**: Salva os dados extraídos em uma planilha do Excel de forma estruturada.
- **Suporte a Grandes Volumes**: Projetado para processar **grandes quantidades de arquivos e dados** sem comprometer o desempenho.

---

## Pré-requisitos

- Python 3.x
- Bibliotecas Python:
  - `os`
  - `glob`
  - `logging`
  - `xml.etree.ElementTree`
  - `openpyxl`

---

## Instalação

1. Clone este repositório:
   git clone https://github.com/seu-usuario/seu-repositorio.git


Configuração
Antes de executar o script, certifique-se de configurar os seguintes caminhos no código:

DOWNLOAD_PATH: Diretório onde os arquivos XML estão armazenados.

EXCEL_PATH: Caminho do arquivo Excel onde os dados serão salvos.

SHEET_NAME: Nome da aba da planilha onde os dados serão inseridos.

Uso
Para executar o script, use o seguinte comando:
python nome_do_script.py
O script processará todos os arquivos XML no diretório especificado, independentemente da quantidade, e salvará os dados na planilha do Excel.

Logs
O script utiliza o módulo logging para registrar informações, avisos e erros durante a execução. Os logs são exibidos no console e incluem detalhes sobre o processamento de cada arquivo XML, bem como quaisquer erros ou avisos encontrados. Isso facilita a identificação de problemas, especialmente ao lidar com grandes volumes de dados.

Exemplo de Saída
A planilha do Excel será preenchida com duas colunas:

Numero: Número da Nota Fiscal.
Chave: Chave de acesso da Nota Fiscal.
O script é capaz de processar milhares de arquivos e gerar uma planilha organizada e livre de duplicatas.

Desempenho com Grandes Volumes
O script foi projetado para ser eficiente mesmo ao processar quantidades massivas de arquivos e dados. Ele utiliza técnicas como:

Leitura e processamento de arquivos em lote.
Verificação de duplicatas em tempo real.
Gerenciamento de memória otimizado para evitar sobrecarga.
Isso garante que o script possa ser usado em cenários com milhares ou até dezenas de milhares de arquivos XML.

Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests para melhorar este projeto, especialmente em relação ao desempenho com grandes volumes de dados.
