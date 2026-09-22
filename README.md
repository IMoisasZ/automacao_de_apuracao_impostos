# Sistema de Apuração Fiscal - Empresa vs Contabilidade

Este documento apresenta o guia completo do **Sistema de Apuração Fiscal - Empresa vs Contabilidade**, uma aplicação desenvolvida em Python para automatizar o cruzamento, validação e comparação de livros fiscais (Livro de Entradas da Empresa, Livro de ICMS da Contabilidade e Livro de IPI da Contabilidade).

---

## 📋 Sumário

1. [Bibliotecas Utilizadas](#️-bibliotecas-utilizadas)
2. [Estrutura de Pastas e Arquivos](#-estrutura-de-pastas-e-arquivos)
3. [Configuração da Logotipo e Imagem](#️-configuração-da-logotipo-e-imagem)
4. [Como Usar a Interface Gráfica (Tkinter)](#-como-usar-a-interface-gráfica)
5. [Como Gerar o Executável (.exe)](#-geração-do-executável-exe)
   - [Com Imagem (Modo Padrão)](#com-imagem-modo-padrão)
   - [Sem Imagem (Modo Limpo)](#sem-imagem-modo-limpo)
6. [Processamento via Notebook (Jupyter / .ipynb)](#-processamento-via-jupyter-notebook-apuracao_fiscalipynb)
7. [Autor](#-autor)

---

## 🛠️ Bibliotecas Utilizadas

Para o funcionamento correto do sistema, são necessárias as seguintes bibliotecas Python:

- [Pandas](https://pandas.pydata.org/) — Manipulação, limpeza, agrupamento e tratamento de dados tabulares.
- [OpenPyXL](https://openpyxl.readthedocs.io/) — Leitura, escrita e formatação avançada de arquivos Excel (`.xlsx`), incluindo aplicação de estilos condicionais.
- [Pillow (PIL)](https://python-pillow.org/) — Processamento e renderização de imagens (PNG/JPG) para exibição na interface gráfica.
- [Tkinter](https://docs.python.org/3/library/tkinter.html) — Biblioteca nativa do Python para construção da interface gráfica (janelas, botões e campos de texto).
- [PyInstaller](https://pyinstaller.org/?utm_source=gemini) - Empacotamento da aplicação em um executável autônomo (.exe).
- **Sys** e **Os** — Módulos nativos para gerenciamento de caminhos de arquivos e recursos do sistema operacional.

---

## 📂 Estrutura de Pastas e Arquivos

Para que o sistema reconheça a logotipo e organize os diretórios, mantenha a seguinte estrutura no seu projeto:

```text
seu_projeto/
│
├── app_apuracao.py             # Código fonte principal do sistema
├── apuracao_fiscal.ipynb       # Notebook opcional para execução sem tela
├── README.md                   # Documentação do sistema
└── assets/
    └── logo_empresa.png        # Imagem de amostra da logotipo
```

---

## 🖼️ Configuração da Logotipo e Imagem

- **Com imagem:** Para exibir a logotipo da sua empresa no topo da janela do sistema, substitua o arquivo de amostra localizado em `assets/logo_empresa.png` por uma nova imagem mantendo exatamente o mesmo nome (`logo_empresa.png`).
- **Sem imagem:** Caso prefira não exibir nenhuma imagem no topo da tela, basta alterar a variável correspondente no código para string vazia:
  ```python
  caminho_imagem = ""
  ```
  Dessa forma, a tela se ajustará automaticamente ocultando o espaço da logotipo.

---

## 💻 Como Executar o Sistema

1. Certifique-se de que as dependências estão devidamente instaladas com o comando pip install -r requirements.txt[cite: 3].

2. Execute o arquivo principal de inicialização a partir da raiz do projeto:

```Bash
python app.py
```

## 💻 Como Usar a Interface Gráfica

1. Execute o arquivo do sistema (`app_apuracao.py` ou o executável gerado).
2. Na janela que se abrirá, preencha os campos selecionando os três arquivos de entrada obrigatórios (compatíveis com os formatos **`.xlsx`** ou **`.xls`**):
   - **Livro de Entradas (Empresa):** Relatório de compras/entradas da empresa.
   - **Livro de ICMS (Contabilidade):** Relatório de apuração de ICMS fornecido pela contabilidade.
   - **Livro de IPI (Contabilidade):** Relatório de apuração de IPI fornecido pela contabilidade.
3. Defina o nome do arquivo de saída (por padrão: `apuracao.xlsx`).
4. Clique no botão **"Executar Apuração e Gerar Excel"**.
5. O sistema processará o cruzamento por chave composta (`NF + CFOP`), gerará as abas `apuracao` (com formatação condicional verde para `OK` e vermelha para `Valor Divergente`) e `resumo`, exibindo uma mensagem de sucesso ao finalizar.

---

## ⚡ Geração do Executável (.exe)

Para gerar um pacote executável autônomo e portátil que roda em qualquer computador Windows sem a necessidade de instalar o Python, utilize o [PyInstaller](https://pyinstaller.org/). Abra o terminal na pasta do projeto e execute um dos comandos abaixo:

### Com Imagem (Modo Padrão):

```bash
pyinstaller --noconsole --onefile --add-data "assets/logo_empresa.png;assets" app_apuracao.py
```

### Sem Imagem (Modo Limpo):

Caso tenha definido `caminho_imagem = ""`, gere o executável sem incluir os assets:

```bash
pyinstaller --noconsole --onefile app_apuracao.py
```

O arquivo `.exe` final estará disponível na pasta `dist/` gerada automaticamente.

---

## 📊 Processamento via Jupyter Notebook (`apuracao_fiscal.ipynb`)

Caso deseje rodar o fluxo completo de apuração e cruzamento diretamente por script ou notebook interativo sem precisar da interface gráfica (Tkinter), utilize o arquivo `apuracao_fiscal.ipynb`.

Para isso, certifique-se de criar uma pasta chamada **`dataset`** na raiz do projeto e posicionar os três arquivos de entrada com os seguintes nomes obrigatórios:

- `dataset/livro_empresa.xlsx`
- `dataset/livro_icms_contabilidade.xlsx`
- `dataset/livro_ipi_contabilidade.xlsx`

---

## 👤 Autor

- **Autor:** Moisés Santos
- **GitHub:** [IMoisasZ](https://github.com/IMoisasZ)
- **E-mail:** [mopri08@gmail.com](mailto:mopri08@gmail.com)
