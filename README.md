# 📄 PDF para XLSX - CILIA Converter

Este projeto converte arquivos **PDF extraídos do sistema CILIA** (voltado para oficinas e seguradoras) em **planilhas Excel (.xlsx)**.

O objetivo principal é facilitar a visualização e o trabalho com os dados de peças e serviços, extraindo apenas as informações essenciais:

- Código da peça
- Quantidade
- Nome da peça

---

## 🔧 Funcionalidades

- Leitura de PDFs exportados do sistema CILIA
- Extração precisa dos campos necessários
- Geração de planilhas XLSX organizadas
- Ideal para **oficinas mecânicas** e **seguradoras** que precisam fazer cotações rápidas

---

## 💻 Como Usar

1. Clone o repositório:

```bash
git clone https://github.com/eukevinbruno/pdf_para_xlsx.git
cd pdf_para_xlsx
```

2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Execute o conversor:

```bash
python conversor.py
```

4. O arquivo XLSX será gerado na mesma pasta contendo a lista com os códigos, quantidades e nomes das peças.

---

## 🧠 Observações

- Certifique-se de usar **PDFs que foram gerados diretamente pelo sistema CILIA**.
- Outros formatos ou PDFs digitalizados podem não funcionar corretamente.

---

## 👨‍💻 Autor

Desenvolvido por **Kevin Bruno** com foco em produtividade para profissionais do setor automotivo.

---
