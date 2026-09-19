# %% [markdown]
# # Sistema para apuração de fiscal - Empresa vs Contabilidade

# %% [markdown]
# ## Bibliotecas

# %%
import sys
import os
import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# %% [markdown]
# ## Função para selecionar os arquivos

# %%
def selecionar_arquivo(entry_widget):
    """Abre a janela para escolher o arquivo Excel correspondente"""
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, arquivo)

# %% [markdown]
# ## Função de validação inicial

# %%
def validacao_inicial(caminho_empresa: str, caminho_icms: str, caminho_ipi: str, nome_saida: str):
# Validações iniciais
    if not caminho_empresa or not caminho_icms or not caminho_ipi:
        messagebox.showwarning("Atenção", "Por favor, selecione os três arquivos de entrada!")
        return
    
    if not nome_saida:
        nome_saida = "apuracao.xlsx"
    elif not nome_saida.endswith(".xlsx"):
        nome_saida += ".xlsx"

# %% [markdown]
# ## Função para caminho absoluto da imagem

# %%
def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funcionando tanto no modo script quanto no .exe """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# %% [markdown]
# ## Função para construção da tela

# %%
def construir_tela():
    """Função responsável exclusivamente por desenhar a interface gráfica"""
    cor_fundo = "#f4f6f9"
    caminho_imagem = resource_path("assets/logo_empresa.png")
    root = tk.Tk()
    root.title("Ferramenta de Apuração Fiscal - Contabilidade")
    root.geometry("480x420" if not caminho_imagem else "480x520")
    root.resizable(False, False)
    root.configure(bg=cor_fundo)
    
    try:
        img_original = Image.open(caminho_imagem)
        
        # Opcional: Redimensione a imagem se ela for muito grande (ex: largura 150px, mantendo a proporção)
        img_original = img_original.resize((150, 80), Image.Resampling.LANCZOS)
        
        img_topo = ImageTk.PhotoImage(img_original)
        
        # Criamos o Label guardando a referência da imagem para o Python não apagá-la da memória (garbage collector)
        lbl_imagem = tk.Label(root, image=img_topo, bg=cor_fundo)
        lbl_imagem.image = img_topo  # Referência vital no Tkinter!
        lbl_imagem.pack(pady=(15, 5)) # Espaçamento acima e abaixo da imagem
    except Exception as e:
        # Se a imagem não for encontrada, o sistema avisa no console mas a tela abre normalmente
        print(f"Aviso: Imagem não carregada ({e}). Certifique-se de que o arquivo 'logo.png' está na pasta.")

    tk.Label(root, text="Automação de Cruzamento Fiscal", font=("Arial", 14, "bold"), bg=cor_fundo).pack(pady=15)

    # Campo 1: Empresa
    tk.Label(root, text="Livro de Entradas (Empresa):", anchor="w", font=("Arial", 9, "bold"), bg=cor_fundo).pack(fill="x", padx=25)
    f1 = tk.Frame(root)
    f1.pack(fill="x", padx=25, pady=3)
    entry_empresa = tk.Entry(f1, width=60)
    entry_empresa.pack(side="left", padx=(0, 5))
    tk.Button(f1, text="Procurar", command=lambda: selecionar_arquivo(entry_empresa), bg=cor_fundo).pack(side="left")

    # Campo 2: ICMS - Contabilidade
    tk.Label(root, text="Livro de ICMS (Contabilidade):", anchor="w", font=("Arial", 9, "bold"), bg=cor_fundo).pack(fill="x", padx=25, pady=(8,0))
    f2 = tk.Frame(root)
    f2.pack(fill="x", padx=25, pady=3)
    entry_icms = tk.Entry(f2, width=60)
    entry_icms.pack(side="left", padx=(0, 5))
    tk.Button(f2, text="Procurar", command=lambda: selecionar_arquivo(entry_icms), bg=cor_fundo).pack(side="left")

    # Campo 3: IPI - Contabilidade
    tk.Label(root, text="Livro de IPI (Contabilidade):", anchor="w", font=("Arial", 9, "bold"), bg=cor_fundo).pack(fill="x", padx=25, pady=(8,0))
    f3 = tk.Frame(root)
    f3.pack(fill="x", padx=25, pady=3)
    entry_ipi = tk.Entry(f3, width=60)
    entry_ipi.pack(side="left", padx=(0, 5))
    tk.Button(f3, text="Procurar", command=lambda: selecionar_arquivo(entry_ipi), bg=cor_fundo).pack(side="left")

    # Campo 4: Nome do arquivo final
    tk.Label(root, text="Nome do Relatório de Saída (.xlsx):", anchor="w", font=("Arial", 9, "bold"), bg=cor_fundo).pack(fill="x", padx=25, pady=(8,0))
    entry_saida = tk.Entry(root, width=67)
    entry_saida.pack(padx=25, pady=3, anchor="w")
    entry_saida.insert(0, "apuracao.xlsx")

    # Botão Executar Principal (passando os campos por argumento)
    btn_executar = tk.Button(
        root, 
        text="Executar Apuração e Gerar Excel", 
        bg="#20055f", 
        fg="white", 
        font=("Arial", 11, "bold"), 
        padx=10, 
        pady=5,
        command=lambda: executar_apuracao(entry_empresa, entry_icms, entry_ipi, entry_saida)
    )
    btn_executar.pack(pady=20)

    root.mainloop()

# %% [markdown]
# ## Formatar arquivo excel

# %%
def formatar_arquivo_excel(nome_arquivo):
    """Aplica a formatação condicional de cores (verde/vermelho) no Excel gerado"""
    try:
        wb = load_workbook(nome_arquivo)
        if 'apuracao' in wb.sheetnames:
            ws = wb['apuracao']

            fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            fill_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

            col_idx_status = None
            for col_i, cell in enumerate(ws[1], 1):
                if cell.value == 'status':
                    col_idx_status = col_i
                    break

            if col_idx_status:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx_status)
                    if cell.value == 'OK':
                        cell.fill = fill_verde
                    elif cell.value == 'Valor Divergente':
                        cell.fill = fill_vermelho

            wb.save(nome_arquivo)
    except Exception as e:
        print(f"Erro ao formatar o Excel: {e}")

# %% [markdown]
# ## Função Principal

# %%
def executar_apuracao(entry_empresa, entry_icms, entry_ipi, entry_saida):
    """Executa toda a lógica de leitura, cruzamento e geração dos relatórios"""
    caminho_empresa = entry_empresa.get()
    caminho_icms = entry_icms.get()
    caminho_ipi = entry_ipi.get()
    nome_saida = entry_saida.get().strip()

    # Validações iniciais
    if not caminho_empresa or not caminho_icms or not caminho_ipi:
        messagebox.showwarning("Atenção", "Por favor, selecione os três arquivos de entrada!")
        return
    
    if not nome_saida:
        nome_saida = "apuracao.xlsx"
    elif not nome_saida.endswith(".xlsx"):
        nome_saida += ".xlsx"

    try:
        # 1.0 Dados Empresa
        df_empresa = pd.read_excel(caminho_empresa, header=1)
        df_empresa = df_empresa.rename(columns={
            'Número': 'nf_empresa',
            'Tributado': 'vl_icms_empresa',
            'Tributado.1': 'vl_ipi_empresa',
            'Nat. Op.': 'cfop_empresa',
            'Valor Contábil': 'vl_cont_empresa',
        })

        delete_columns_empresa = [
            'Espécie', 'Série', 'Emissão', 'Entrada', 'Cód Emitente', 'UF', 
            'Incidência', '%', 'Base Cálc.', 'Isento', 'Outros', 
            'Incidência.1', '%.1', 'Base Cálc..1', 
            'Observação Substituição Tributária', 'Isento.1', 'Outros.1', 'Observação'
        ]
        df_empresa = df_empresa.drop(columns=delete_columns_empresa, errors='ignore').dropna()

        df_empresa['nf_empresa'] = pd.to_numeric(df_empresa['nf_empresa'], errors='coerce').astype('Int64')
        df_empresa['cfop_empresa'] = pd.to_numeric(df_empresa['cfop_empresa'], errors='coerce').astype('Int64')
        df_empresa = df_empresa.dropna(subset=['nf_empresa', 'cfop_empresa'])

        df_empresa['id_unico_empresa'] = df_empresa['nf_empresa'].astype(str) + '|' + df_empresa['cfop_empresa'].astype(str)
        df_empresa = df_empresa.groupby('id_unico_empresa', as_index=False).sum(numeric_only=True).fillna(0)

        qtde_nf_empresa = df_empresa['nf_empresa'].dropna().drop_duplicates().count()
        valor_contabil_total_empresa = df_empresa['vl_cont_empresa'].sum()
        valor_icms_total_empresa = df_empresa['vl_icms_empresa'].sum()
        valor_ipi_total_empresa = df_empresa['vl_ipi_empresa'].sum()


        # 2.0 Dados Contabilidade - ICMS
        df_contabilidade_icms = pd.read_excel(caminho_icms, header=6)
        df_contabilidade_icms = df_contabilidade_icms.rename(columns={
            'Vlr Contábil': 'vl_cont_icms_contabilidade',
            'Valor ICMS': 'vl_icms_contabilidade',
            'Número': 'nf_icms_contabilidade',
            'Natureza': 'cfop_icms'
        })
        df_contabilidade_icms = df_contabilidade_icms.drop(columns=[
            'Nr Lcto', 'Pessoa', 'Data Lcto', 'Esp', 'Série', 
            'Base ICMS', 'Isentas ICMS', 'Outras ICMS'
        ], errors='ignore').dropna()

        df_contabilidade_icms = df_contabilidade_icms[df_contabilidade_icms['cfop_icms'] != 'Natureza']
        df_contabilidade_icms['cfop_icms'] = pd.to_numeric(df_contabilidade_icms['cfop_icms'], errors='coerce')
        df_contabilidade_icms['nf_icms_contabilidade'] = pd.to_numeric(df_contabilidade_icms['nf_icms_contabilidade'], errors='coerce')
        df_contabilidade_icms = df_contabilidade_icms.dropna(subset=['cfop_icms', 'nf_icms_contabilidade'])

        df_contabilidade_icms['cfop_icms'] = df_contabilidade_icms['cfop_icms'].astype('int')
        df_contabilidade_icms['nf_icms_contabilidade'] = df_contabilidade_icms['nf_icms_contabilidade'].astype('int')

        for col in ['vl_icms_contabilidade', 'vl_cont_icms_contabilidade']:
            df_contabilidade_icms[col] = pd.to_numeric(df_contabilidade_icms[col], errors='coerce').fillna(0)

        df_contabilidade_icms = df_contabilidade_icms.query('cfop_icms <= 4000').copy()
        df_contabilidade_icms['id_unico_icms'] = df_contabilidade_icms['nf_icms_contabilidade'].astype(str) + '|' + df_contabilidade_icms['cfop_icms'].astype(str)
        df_contabilidade_icms = df_contabilidade_icms.groupby('id_unico_icms', as_index=False).sum(numeric_only=True)

        qtde_nf_icms_contabilidade = df_contabilidade_icms['nf_icms_contabilidade'].dropna().drop_duplicates().count()
        valor_contabil_total_icms_contabilidade = df_contabilidade_icms['vl_cont_icms_contabilidade'].sum()
        valor_icms_total_contabilidade = df_contabilidade_icms['vl_icms_contabilidade'].sum() 


        # 3.0 Dados Contabilidade - IPI
        df_contabilidade_ipi = pd.read_excel(caminho_ipi, header=6)
        df_contabilidade_ipi = df_contabilidade_ipi.rename(columns={
            'Vlr Contabil': 'vl_cont_ipi_contabilidade',
            'Valor IPI': 'vl_ipi_contabilidade',
            'Número': 'nf_ipi_contabilidade',
            'Natureza': 'cfop_ipi'
        })
        df_contabilidade_ipi = df_contabilidade_ipi.drop(columns=[
            'Nr Lcto', 'Fornecedor', 'Data Lcto', 'Esp', 'Série', 
            'Base IPI', 'Isentas IPI', 'Outras IPI'
        ], errors='ignore').dropna()

        df_contabilidade_ipi = df_contabilidade_ipi[df_contabilidade_ipi['cfop_ipi'] != 'Natureza']
        df_contabilidade_ipi['cfop_ipi'] = pd.to_numeric(df_contabilidade_ipi['cfop_ipi'], errors='coerce')
        df_contabilidade_ipi['nf_ipi_contabilidade'] = pd.to_numeric(df_contabilidade_ipi['nf_ipi_contabilidade'], errors='coerce')
        df_contabilidade_ipi = df_contabilidade_ipi.dropna(subset=['cfop_ipi', 'nf_ipi_contabilidade'])

        df_contabilidade_ipi['cfop_ipi'] = df_contabilidade_ipi['cfop_ipi'].astype('int')
        df_contabilidade_ipi['nf_ipi_contabilidade'] = df_contabilidade_ipi['nf_ipi_contabilidade'].astype('int')

        for col in ['vl_ipi_contabilidade', 'vl_cont_ipi_contabilidade']:
            df_contabilidade_ipi[col] = pd.to_numeric(df_contabilidade_ipi[col], errors='coerce').fillna(0)

        df_contabilidade_ipi = df_contabilidade_ipi.query('cfop_ipi <= 4000').copy()
        df_contabilidade_ipi['id_unico_ipi'] = df_contabilidade_ipi['nf_ipi_contabilidade'].astype(str) + '|' + df_contabilidade_ipi['cfop_ipi'].astype(str)
        df_contabilidade_ipi = df_contabilidade_ipi.groupby('id_unico_ipi', as_index=False).sum(numeric_only=True)

        valor_ipi_total_contabilidade = df_contabilidade_ipi['vl_ipi_contabilidade'].sum()


        # 4.0 Lista Mestre e Cruzamentos
        lista_mestre = pd.concat([
            df_empresa['id_unico_empresa'],
            df_contabilidade_icms['id_unico_icms'],
            df_contabilidade_ipi['id_unico_ipi']
        ]).dropna().drop_duplicates()

        df_mestre = pd.DataFrame({'id_unico': lista_mestre})
        df_mestre[['nf', 'cfop']] = df_mestre['id_unico'].str.split('|', expand=True)

        # 5.0 União das Bases (Left Join)
        df_mestre_e_empresa = pd.merge(df_mestre, df_empresa, left_on='id_unico', right_on='id_unico_empresa', how='left')
        df_mestre_e_empresa_e_icms = pd.merge(df_mestre_e_empresa, df_contabilidade_icms, left_on='id_unico', right_on='id_unico_icms', how='left')
        df_unido = pd.merge(df_mestre_e_empresa_e_icms, df_contabilidade_ipi, left_on='id_unico', right_on='id_unico_ipi', how='left')

        df_unido['cfop'] = pd.to_numeric(df_unido['cfop'], errors='coerce').astype('Int64')
        df_unido['nf'] = pd.to_numeric(df_unido['nf'], errors='coerce').astype('Int64')
        df_unido['nf_icms_contabilidade'] = pd.to_numeric(df_unido['nf_icms_contabilidade'], errors='coerce').astype('Int64')
        df_unido['cfop_icms'] = pd.to_numeric(df_unido['cfop_icms'], errors='coerce').astype('Int64')

        df_unido = df_unido.fillna(0)
        df_unido['dif_vl_contabil'] = (df_unido['vl_cont_empresa'] - df_unido['vl_cont_icms_contabilidade']).round(2)
        df_unido['dif_vl_icms'] = (df_unido['vl_icms_empresa'] - df_unido['vl_icms_contabilidade']).round(2)
        df_unido['dif_vl_ipi'] = (df_unido['vl_ipi_empresa'] - df_unido['vl_ipi_contabilidade']).round(2)

        cols_dif = ['dif_vl_contabil', 'dif_vl_icms', 'dif_vl_ipi']
        df_unido['status'] = 'Valor Divergente'
        df_unido.loc[df_unido[cols_dif].eq(0).all(axis=1), 'status'] = 'OK'
        df_unido = df_unido.sort_values(by=['nf', 'cfop'])

        cols_desejadas = ['nf', 'cfop', 'vl_cont_empresa', 'vl_icms_empresa', 'vl_ipi_empresa', 'vl_cont_icms_contabilidade', 'vl_icms_contabilidade', 'vl_ipi_contabilidade', 'dif_vl_contabil', 'dif_vl_icms', 'dif_vl_ipi', 'status']
        cols_existentes = [col for col in cols_desejadas if col in df_unido.columns]
        df_export_excel = df_unido[cols_existentes]

        # 6.0 DataFrame Resumo
        df_resumo = pd.DataFrame({
            'qtde_nfs_empresa': [qtde_nf_empresa],
            'qtde_nfs_contabilidade': [qtde_nf_icms_contabilidade],
            'dif_nfs': ([qtde_nf_empresa - qtde_nf_icms_contabilidade]),
            'vl_cont_total_empresa': [valor_contabil_total_empresa],
            'vl_cont_total_icms_contabilidade': [valor_contabil_total_icms_contabilidade],
            'dif_vl_cont_total': ([valor_contabil_total_empresa - valor_contabil_total_icms_contabilidade]),
            'vl_total_icms_empresa': [valor_icms_total_empresa],
            'vl_total_icms_contabilidade': [valor_icms_total_contabilidade],
            'dif_vl_total_icms': ([valor_icms_total_empresa - valor_icms_total_contabilidade]),
            'vl_total_ipi_empresa': [valor_ipi_total_empresa],
            'vl_total_ipi_contabilidade': [valor_ipi_total_contabilidade],
            'dif_vl_total_ipi': ([valor_ipi_total_empresa - valor_ipi_total_contabilidade])
        })
        
        cols_formatar = [
            'vl_cont_total_empresa',
            'vl_cont_total_icms_contabilidade',
            'dif_vl_cont_total',
            'vl_total_icms_empresa',
            'vl_total_icms_contabilidade',
            'dif_vl_total_icms',
            'vl_total_ipi_empresa',
            'vl_total_ipi_contabilidade',
            'dif_vl_total_ipi']

        for col in cols_formatar:
            df_resumo[col] = df_resumo[col].round(2)

        # Salvando as abas no Excel
        with pd.ExcelWriter(nome_saida, engine='openpyxl') as writer:
            df_export_excel.to_excel(writer, sheet_name='apuracao', index=False)
            df_resumo.to_excel(writer, sheet_name='resumo', index=False)

        # Chama a função separada para formatar as cores da planilha
        formatar_arquivo_excel(nome_saida)

        messagebox.showinfo("Sucesso!", f"Apuração concluída e salva com sucesso em:\n{nome_saida}")

    except Exception as e:
        messagebox.showerror("Erro no Processamento", f"Ocorreu um erro durante a execução:\n{str(e)}")
        
if __name__ == "__main__":
    construir_tela()


