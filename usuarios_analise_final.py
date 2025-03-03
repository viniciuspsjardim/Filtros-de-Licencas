import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import matplotlib.pyplot as plt
import csv
from fpdf import FPDF


def carregar_csv():
    arquivo = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if arquivo:
        global df
        df = pd.read_csv(arquivo)
        messagebox.showinfo("Arquivo carregado", f"{arquivo} carregado com sucesso!")
        atualizar_tabela(df)
        atualizar_resumo(df)
        atualizar_licenciados()
        
def exportar_csv():
    if not df_filtrado.empty:
        arquivo = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if arquivo:
            df_filtrado.to_csv(arquivo, index=False)
            messagebox.showinfo("Exportação CSV", "Arquivo CSV exportado com sucesso!")


def exportar_pdf():
    if not df_filtrado.empty:
        arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if arquivo:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)

            colunas = ['Nome para exibição', 'Nome UPN', 'Licenças', 'Bloquear credencial']
            for col in colunas:
                pdf.cell(48, 8, col, border=1)
            pdf.ln()

            for _, row in df_filtrado.iterrows():
                pdf.cell(48, 8, str(row['Nome para exibição']), border=1)
                pdf.cell(48, 8, str(row['Nome UPN']), border=1)
                pdf.cell(48, 8, str(row['Licenças']), border=1)
                pdf.cell(48, 8, str(row['Bloquear credencial']), border=1)
                pdf.ln()

            pdf.output(arquivo)
            messagebox.showinfo("Exportação PDF", "Arquivo PDF exportado com sucesso!")


def atualizar_tabela(dados):
    for row in tree_usuarios.get_children():
        tree_usuarios.delete(row)
    for _, row in dados.iterrows():
        tree_usuarios.insert("", tk.END, values=(row['Nome para exibição'], row['Nome UPN'], row['Licenças'], row['Bloquear credencial']))


def atualizar_resumo(dados):
    for row in tree_resumo.get_children():
        tree_resumo.delete(row)

    empresas = ['autonomoz', 'bellavia', 'rhyno', 'hbko', '#EXT#']
    total_ativos = 0
    total_inativos = 0
    total_geral = 0

    for empresa in empresas:
        filtro = dados[dados['Nome UPN'].str.contains(empresa, case=False, na=False)]
        ativos = len(filtro[filtro['Bloquear credencial'] == False])
        inativos = len(filtro[filtro['Bloquear credencial'] == True])
        total = ativos + inativos
        total_ativos += ativos
        total_inativos += inativos
        total_geral += total
        tree_resumo.insert("", tk.END, values=(empresa.capitalize(), ativos, inativos, total))

    # Adicionando a linha de total
    tree_resumo.insert("", tk.END, values=("TOTAL", total_ativos, total_inativos, total_geral), tags=('total',))
    tree_resumo.tag_configure('total', background='#d9d9d9', font=('Arial', 10, 'bold'))


def exibir_grafico():
    empresas = ['autonomoz', 'bellavia', 'rhyno', 'hbko', '#EXT#']
    ativos = []
    inativos = []

    for empresa in empresas:
        filtro = df[df['Nome UPN'].str.contains(empresa, case=False, na=False)]
        ativos.append(len(filtro[filtro['Bloquear credencial'] == False]))
        inativos.append(len(filtro[filtro['Bloquear credencial'] == True]))

    x = range(len(empresas))

    plt.figure(figsize=(8, 5))
    plt.bar(x, ativos, width=0.4, label='Ativos', align='center')
    plt.bar(x, inativos, width=0.4, bottom=ativos, label='Inativos', align='center')

    for i in x:
        plt.text(i, ativos[i]/2, str(ativos[i]), ha='center', va='center', color='white', fontsize=10)
        plt.text(i, ativos[i] + inativos[i]/2, str(inativos[i]), ha='center', va='center', color='white', fontsize=10)

    total_usuarios = len(df)
    plt.xticks(x, ['Autonomoz', 'Bellavia', 'Rhyno', 'HBKO', 'EXT'])
    plt.ylabel('Quantidade')
    plt.title(f'Usuários Ativos/Inativos por Empresa (Total: {total_usuarios})')
    plt.legend()
    plt.show()


def atualizar_licenciados():
    for row in tree_licenciados.get_children():
        tree_licenciados.delete(row)

    if df is None:
        return

    # Licenças selecionadas
    licencas_filtro = []
    if var_basic.get():
        licencas_filtro.append("Microsoft 365 Business Basic")
    if var_standard.get():
        licencas_filtro.append("Microsoft 365 Business Standard")
    if var_exchange.get():
        licencas_filtro.append("Exchange Online (Plan 2)")
    if var_powerbi.get():
        licencas_filtro.append("Power BI Pro")

    if not licencas_filtro:
        return  # Se nenhuma licença estiver marcada, não exibe nada

    # Empresas selecionadas
    empresas = []
    if var_autonomoz.get():
        empresas.append("autonomoz")
    if var_bellavia.get():
        empresas.append("bellavia")
    if var_rhyno.get():
        empresas.append("rhyno")
    if var_hbko.get():
        empresas.append("hbko")

    # Filtrar dados com as licenças selecionadas
    filtro = df[df['Licenças'].str.contains('|'.join(licencas_filtro), case=False, na=False)]

    if empresas:
        filtro = filtro[filtro['Nome UPN'].str.contains('|'.join(empresas), case=False, na=False)]

    
    global df_filtrado
    df_filtrado = pd.DataFrame(columns=['Nome para exibição', 'Nome UPN', 'Licenças', 'Bloquear credencial'])
    # Preencher a tabela exibindo apenas as licenças selecionadas
    for _, row in filtro.iterrows():
        licencas_usuario = [licenca for licenca in licencas_filtro if licenca.lower() in row['Licenças'].lower()]
        if not licencas_usuario:
            continue

        licencas_exibidas = ', '.join(licencas_usuario)
        valores = (row['Nome para exibição'], row['Nome UPN'], licencas_exibidas, row['Bloquear credencial'])
        item = tree_licenciados.insert("", tk.END, values=valores)
        
        if row['Bloquear credencial']:
            tree_licenciados.item(item, tags=('inativo',))
        
        # Adiciona no DataFrame para exportação
        df_filtrado.loc[len(df_filtrado)] = [row['Nome para exibição'], row['Nome UPN'], licencas_exibidas, row['Bloquear credencial']]





# Interface gráfica
df = None
root = tk.Tk()
root.title("Análise de Usuários por Empresa")
root.iconbitmap('icon.ico')

frame_topo = tk.Frame(root)
frame_topo.pack(padx=10, pady=10)
btn_carregar = tk.Button(frame_topo, text="Carregar CSV", command=carregar_csv)
btn_carregar.pack()

notebook = ttk.Notebook(root)
notebook.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Aba Usuários
aba_usuarios = tk.Frame(notebook)
notebook.add(aba_usuarios, text="Usuários")

tk.Label(aba_usuarios, text="Usuários").pack()
colunas_usuarios = ('Nome para exibição', 'Nome UPN', 'Licenças', 'Bloquear credencial')
tree_usuarios = ttk.Treeview(aba_usuarios, columns=colunas_usuarios, show='headings')
for col in colunas_usuarios:
    tree_usuarios.heading(col, text=col)
    tree_usuarios.column(col, width=150)
tree_usuarios.pack(fill=tk.BOTH, expand=True)

# Aba Usuários Ativos por Empresa
aba_resumo = tk.Frame(notebook)
notebook.add(aba_resumo, text="Usuários Ativos por Empresa")

tk.Label(aba_resumo, text="Usuários Ativos por Empresa").pack()
colunas_resumo = ('Empresa', 'Ativos', 'Inativos', 'Total')
tree_resumo = ttk.Treeview(aba_resumo, columns=colunas_resumo, show='headings')
for col in colunas_resumo:
    tree_resumo.heading(col, text=col)
    tree_resumo.column(col, width=100)
tree_resumo.pack(fill=tk.BOTH, expand=True)

btn_grafico = tk.Button(aba_resumo, text="Exibir Gráfico de Usuários", command=exibir_grafico)
btn_grafico.pack(pady=10)

# Aba Usuários Licenciados
aba_licenciados = tk.Frame(notebook)
notebook.add(aba_licenciados, text="Usuários Licenciados")

tk.Label(aba_licenciados, text="Usuários Licenciados").pack()
var_autonomoz = tk.BooleanVar()
var_bellavia = tk.BooleanVar()
var_rhyno = tk.BooleanVar()
var_hbko = tk.BooleanVar()

frame_filtros = tk.Frame(aba_licenciados)
frame_filtros.pack()

tk.Checkbutton(frame_filtros, text="Autonomoz", variable=var_autonomoz, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros, text="Bellavia", variable=var_bellavia, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros, text="Rhyno", variable=var_rhyno, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros, text="HBKO", variable=var_hbko, command=atualizar_licenciados).pack(side=tk.LEFT)

# Variáveis para as licenças
var_basic = tk.BooleanVar(value=True)
var_standard = tk.BooleanVar(value=True)
var_exchange = tk.BooleanVar(value=True)
var_powerbi = tk.BooleanVar(value=True)

frame_filtros_licencas = tk.Frame(aba_licenciados)
frame_filtros_licencas.pack(pady=5)

tk.Label(frame_filtros_licencas, text="Filtrar Licenças:").pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros_licencas, text="Basic", variable=var_basic, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros_licencas, text="Standard", variable=var_standard, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros_licencas, text="Exchange", variable=var_exchange, command=atualizar_licenciados).pack(side=tk.LEFT)
tk.Checkbutton(frame_filtros_licencas, text="Power BI", variable=var_powerbi, command=atualizar_licenciados).pack(side=tk.LEFT)

colunas_licenciados = ('Nome para exibição', 'Nome UPN', 'Licenças', 'Bloquear credencial')
tree_licenciados = ttk.Treeview(aba_licenciados, columns=colunas_licenciados, show='headings')
for col in colunas_licenciados:
    tree_licenciados.heading(col, text=col)
    tree_licenciados.column(col, width=150)
tree_licenciados.tag_configure('inativo', background='red')
tree_licenciados.pack(fill=tk.BOTH, expand=True)

frame_botoes_exportar = tk.Frame(aba_licenciados)
frame_botoes_exportar.pack(pady=10)

btn_exportar_csv = tk.Button(frame_botoes_exportar, text="Exportar CSV", command=exportar_csv)
btn_exportar_csv.pack(side=tk.LEFT, padx=5)

btn_exportar_pdf = tk.Button(frame_botoes_exportar, text="Exportar PDF", command=exportar_pdf)
btn_exportar_pdf.pack(side=tk.LEFT, padx=5)

root.mainloop()
