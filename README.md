
---

## README.md

```markdown
# Extrator Web Inteligente de PDF para Planilhas Excel

Uma aplicação web desenvolvida com Python e Flask que permite aos usuários fazer upload de arquivos PDF, extrair dados tabulares específicos (QTD, CÓDIGO, TITULO) e baixar os resultados como planilhas Excel (.xlsx). A ferramenta é capaz de processar múltiplos PDFs, lidar com títulos que se estendem por várias linhas e agrupar múltiplos resultados em um arquivo ZIP.

## Visão Geral

Este projeto foi criado para automatizar o processo, muitas vezes tedioso e propenso a erros, de extrair manualmente informações de tabelas contidas em documentos PDF. A aplicação utiliza a biblioteca `pdfplumber` para uma análise detalhada do conteúdo do PDF, incluindo a posição de cada palavra, permitindo uma extração robusta mesmo em layouts complexos onde a detecção automática de tabelas pode falhar.

## Funcionalidades Principais

* **Interface Web Amigável:** Upload de um ou mais arquivos PDF através de uma interface estilizada com Tailwind CSS.
* **Extração Seletiva de Dados:** Foco na extração das colunas "QTD", "CÓDIGO" e "TITULO".
* **Lógica de Extração Avançada:** Análise baseada em texto e coordenadas para identificar e extrair dados.
* **Tratamento de Títulos Multilinha:** Algoritmo customizado para agrupar corretamente segmentos de títulos que se estendem por múltiplas linhas visuais.
* **Saída em Formato Excel:** Gera arquivos `.xlsx` para cada PDF.
* **Download Consolidado:** Se múltiplos PDFs são processados, os arquivos Excel resultantes são agrupados em um único arquivo `.zip`.
* **Gerenciamento de Arquivos Temporários:** PDFs enviados e arquivos Excel/ZIP gerados são automaticamente removidos do servidor após o processamento.
* **Acesso em Rede Local:** Pode ser configurado para acesso na rede local através de um nome de domínio personalizado (via edição do arquivo `hosts` e configuração do Nginx como proxy reverso, se desejado).

## Tecnologias Utilizadas

* **Back-end:** Python 3, Flask
* **Extração de PDF:** `pdfplumber`
* **Manipulação de Dados:** `pandas`
* **Geração de Excel:** `openpyxl` (como dependência do pandas para escrita em `.xlsx`)
* **Front-end:** HTML, Tailwind CSS (via CDN)
* **Utilitários:** `zipfile` (para criar arquivos .zip), `os`, `uuid`

## Configuração e Execução

### Pré-requisitos

* Python 3.7 ou superior.
* `pip` (gerenciador de pacotes Python).

### Instalação de Dependências

1.  Clone este repositório (ou crie a estrutura de arquivos descrita acima e adicione `app.py`).
2.  Navegue até o diretório raiz do projeto no seu terminal.
3.  Crie e ative um ambiente virtual (recomendado):
    ```bash
    python -m venv venv
    # Windows
    .\venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```
4.  Instale as dependências listadas no arquivo `requirements.txt`:
    ```bash
    pip install Flask pandas pdfplumber Pillow openpyxl
    ```
    (Você pode criar um arquivo `requirements.txt` com estas dependências listadas, uma por linha).

### Executando a Aplicação

1.  No terminal, com o ambiente virtual ativado e no diretório raiz do projeto, execute:
    ```bash
    python app.py
    ```
2.  A aplicação Flask será iniciada. Por padrão (com `debug=True` e `host='0.0.0.0'`), ela estará acessível em:
    * `http://127.0.0.1:5000/` (ou `http://localhost:5000/`)
    * E também através do IP da sua máquina na rede local (ex: `http://192.168.1.100:5000/`).
3.  Abra o endereço no seu navegador web para acessar a interface de upload.

### Configuração para Acesso com Nome de Domínio Local (Opcional)

Para acessar a aplicação usando um nome de domínio personalizado na sua rede local (ex: `http://meu-extrator.local` em vez de `http://IP_DA_MAQUINA:5000`), você pode:

1.  **Editar o arquivo `hosts`:**
    * **Windows:** `C:\Windows\System32\drivers\etc\hosts`
    * **macOS/Linux:** `/etc/hosts`
    * Adicione uma entrada como: `SEU_IP_LOCAL meu-extrator.local`
2.  **Usar um Proxy Reverso (Recomendado para múltiplos "sites" na porta 80):**
    * Configure um servidor web como Nginx ou Apache para escutar na porta 80 e redirecionar o tráfego para a porta da sua aplicação Flask (ex: 5000) com base no nome do host. Isso permite que você acesse `http://meu-extrator.local` sem especificar a porta.

## Estrutura do Código Principal (`app.py`)

O arquivo `app.py` contém:

1.  **Classe `ExtratorTabelaPDF`:** Encapsula toda a lógica de parsing do PDF, identificação de cabeçalhos, extração de dados linha a linha, e tratamento de títulos multilinha.
    * `_encontrar_limites_colunas_cabecalho()`: Método interno para definir as coordenadas das colunas.
    * `_extrair_dados_baseado_em_texto()`: Método interno que realiza a extração principal.
    * `processar_pdf()`: Método público para processar um arquivo PDF.
    * `salvar_resultado_excel()`: Método público para salvar o DataFrame em um arquivo Excel.
2.  **Configuração da Aplicação Flask:**
    * Inicialização do Flask.
    * Definição de pastas para uploads e arquivos de saída temporários.
    * Template HTML embutido para a interface de usuário.
3.  **Rotas Flask:**
    * `@app.route('/', methods=['GET', 'POST'])`: Rota principal para upload de arquivos e download dos resultados.
    * Lógica para lidar com uploads de múltiplos arquivos, processamento, geração de Excel/ZIP, e envio para o usuário.
4.  **Gerenciamento de Arquivos Temporários:**
    * Uso de `flask.g` e `@app.after_request` para garantir a remoção dos arquivos PDF enviados, Excels gerados e ZIPs após cada requisição.

## Possíveis Melhorias Futuras

* Interface para o usuário definir dinamicamente as colunas a serem extraídas.
* Suporte a OCR para PDFs baseados em imagem.
* Melhorias na heurística de detecção de tabelas para PDFs ainda mais variados.
* Paginação ou processamento assíncrono para arquivos PDF muito grandes ou um grande número de uploads simultâneos.
* Autenticação de usuário, se a ferramenta for compartilhada de forma mais ampla.

---
