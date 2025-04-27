import tkinter as tk
from tkinter import messagebox, filedialog, ttk, scrolledtext
import pandas as pd
import json
import os
import re
import html
import webbrowser
import chardet  # Certifique-se de importar a biblioteca

# Dados iniciais
df = pd.DataFrame()
entradas = {}
check_vars = {}
arvore = None
ordenar_crescente = False
coluna_ordenada = None
pagina_atual = 0
linhas_por_pagina = 20

# Dicionário para armazenar o estado dos checkboxes por ID
checkbox_estado = {}

# Arquivo de estado dos checkboxes
ESTADO_CHECKBOX_ARQUIVO = "checkbox_estado.json"

# Salva estado dos checkboxes
def salvar_estado_checkboxes():
    estado = {col: var.get() for col, var in check_vars.items()}
    with open(ESTADO_CHECKBOX_ARQUIVO, "w") as f:
        json.dump(estado, f)

# Carrega estado dos checkboxes
def carregar_estado_checkboxes():
    if os.path.exists(ESTADO_CHECKBOX_ARQUIVO):
        with open(ESTADO_CHECKBOX_ARQUIVO, "r") as f:
            return json.load(f)
    return {}

# Remove tags e entidades HTML
def limpar_html(texto):
    sem_tags = re.sub(r'<[^>]+>', '', str(texto))
    return html.unescape(sem_tags)

# Validação inline
def validar_campos():
    faltantes = []
    for col, widget in entradas.items():
        if check_vars[col].get():
            valor = widget.get("1.0", tk.END).strip() if isinstance(widget, scrolledtext.ScrolledText) else widget.get().strip()
            if not valor:
                faltantes.append(col)
                widget.config(background='misty rose')
            else:
                widget.config(background='white')
    if faltantes:
        messagebox.showwarning("Validação", f"Preencha os campos obrigatórios: {', '.join(faltantes)}")
        return False
    return True

# Padronizar colunas SIM/NÃO
def padronizar_colunas_sim_nao():
    colunas_sim_nao = ['Exibir na loja', 'Frete gratis', 'Produto Físico']
    for col in colunas_sim_nao:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: 'SIM' if str(x).strip().lower() in ['sim', 's', '1'] else 'NÃO')

# Menu de contexto

def copiar_valor():
    sel = arvore.focus()
    if not sel: return
    col_id = arvore.identify_column(arvore.winfo_pointerx() - arvore.winfo_rootx())
    if col_id == '#0': return
    idx = int(col_id.replace('#','')) - 1
    val = arvore.item(sel, 'values')[idx]
    janela.clipboard_clear()
    janela.clipboard_append(val)

def abrir_link():
    sel = arvore.focus()
    if not sel: return
    for v in arvore.item(sel, 'values'):
        if isinstance(v, str) and v.startswith('http'):
            webbrowser.open(v)
            break

def criar_menu_contexto():
    global menu_contexto
    menu_contexto = tk.Menu(janela, tearoff=0)
    menu_contexto.add_command(label='Copiar valor', command=copiar_valor)
    menu_contexto.add_command(label='Abrir link', command=abrir_link)

# Funções auxiliares
def toggle_campo(col):
    w = entradas[col]
    st = 'normal' if check_vars[col].get() else 'disabled'
    w.config(state=st)
    salvar_estado_checkboxes()

def limpar_campos(preservar_estado=False):
    for col,w in entradas.items():
        if isinstance(w, scrolledtext.ScrolledText):
            w.config(state='normal'); w.delete('1.0',tk.END)
        else:
            w.config(state='normal'); w.delete(0,tk.END)
        if not preservar_estado and col in check_vars:
            check_vars[col].set(True)
        toggle_campo(col)

def preencher_campos(evt=None):
    sel = arvore.focus()
    if sel:
        vals = arvore.item(sel,'values')
        cols = df.columns.tolist()
        for i,col in enumerate(cols):
            if col in entradas:
                w = entradas[col]
                txt = limpar_html(vals[i])
                if isinstance(w,scrolledtext.ScrolledText):
                    w.config(state='normal'); w.delete('1.0',tk.END); w.insert(tk.END,txt)
                else:
                    w.config(state='normal'); w.delete(0,tk.END); w.insert(0,txt)
                check_vars[col].set(True)

def aplicar_filtro():
    texto = filter_var.get().lower()
    coluna = coluna_filtro_var.get()  # Obtém a coluna selecionada

    if arvore:
        for item in arvore.get_children():
            arvore.delete(item)

        if texto == "":
            data = df.iterrows()
        elif coluna in df.columns:
            # Filtra apenas pela coluna selecionada
            data = [(i, row) for i, row in df.iterrows() if texto in str(row[coluna]).lower()]
        else:
            messagebox.showwarning("Aviso", f"A coluna '{coluna}' não existe no DataFrame.")
            return

        for i, row in data:
            tag = 'oddrow' if i % 2 else 'evenrow'
            arvore.insert("", "end", values=list(row), tags=(tag,))

def focar_filtro(event=None):
    entry_filtro.focus_set()

def alternar_checkbox(event):
    item = arvore.identify_row(event.y)
    if item:
        valores = list(arvore.item(item, 'values'))
        id_val = arvore.item(item, 'values')[1]  # ID mostrado na tabela
        try:
            index_df = int(float(id_val))
        except (ValueError, TypeError):
            # Ignora se não for um número válido
            return
        # Alterna o valor do checkbox
        if valores[0] == '1':
            valores[0] = ''
            checkbox_estado[index_df] = False
        else:
            valores[0] = '1'
            checkbox_estado[index_df] = True
        arvore.item(item, values=valores)    

# CRUD principal
def adicionar_linha():
    if not validar_campos(): return
    linha={col:(w.get('1.0',tk.END).strip() if isinstance(w,scrolledtext.ScrolledText) else w.get().strip()) for col,w in entradas.items()}
    global df
    df=pd.concat([df,pd.DataFrame([linha])], ignore_index=True)
    atualizar_tabela(selecionar_ultimo=True)

def alterar_linha():
    if not validar_campos(): 
        return

    if alterar_todos_var.get():  # Verifica se o checkbox "Alterar Todos" está marcado
        # Alterar todos os produtos
        for idx in range(len(df)):
            for col, w in entradas.items():
                val = w.get('1.0', tk.END).strip() if isinstance(w, scrolledtext.ScrolledText) else w.get().strip()
                df.at[idx, col] = val  # Atualiza o DataFrame com os novos valores
    else:
        # Alterar apenas os produtos selecionados
        selecionados = arvore.selection()  # Obtém os itens selecionados na tabela
        if selecionados:
            for sel in selecionados:
                idx = arvore.index(sel)  # Obtém o índice da linha selecionada
                for col, w in entradas.items():
                    val = w.get('1.0', tk.END).strip() if isinstance(w, scrolledtext.ScrolledText) else w.get().strip()
                    df.at[idx, col] = val  # Atualiza o DataFrame com os novos valores
        else:
            # Exibe uma mensagem de aviso se nenhum produto estiver selecionado
            messagebox.showwarning("Aviso", "Nenhum produto selecionado para alterar.")

    atualizar_tabela()  # Atualiza a tabela exibida

def detectar_codificacao(arquivo):
    with open(arquivo, 'rb') as f:
        resultado = chardet.detect(f.read())
        return resultado['encoding']

def carregar_csv():
    global df
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if path:
        try:
            # Detectar a codificação do arquivo
            encoding_detectada = detectar_codificacao(path)
            # Carregar o CSV com a codificação detectada
            df = pd.read_csv(path, encoding=encoding_detectada, sep=';')
            padronizar_colunas_sim_nao()
            criar_campos()
            atualizar_tabela()

            # Atualizar o Combobox com as colunas do DataFrame
            coluna_filtro['values'] = df.columns.tolist()
            if df.columns.tolist():
                coluna_filtro.current(0)  # Seleciona a primeira coluna por padrão

            if arvore.get_children():
                primeiro_item = arvore.get_children()[0]
                arvore.selection_set(primeiro_item)
                arvore.focus(primeiro_item)
                preencher_campos()
        except UnicodeDecodeError as e:
            messagebox.showerror("Erro de Codificação", f"Erro ao carregar o arquivo: {e}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

def salvar_csv():
    path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV', '*.csv')])
    if path:
        df.to_csv(path, index=False, encoding='utf-8', sep=';')  # Certifique-se de usar utf-8

def inverter_ordem():
    global ordenar_crescente
    ordenar_crescente=not ordenar_crescente
    atualizar_tabela()

def ordenar_por_coluna(col):
    global df, ordenar_crescente, coluna_ordenada
    if coluna_ordenada == col:
        ordenar_crescente = not ordenar_crescente
    else:
        ordenar_crescente = True
        coluna_ordenada = col
    df.sort_values(by=col, ascending=ordenar_crescente, inplace=True, ignore_index=True)
    atualizar_tabela()

def criar_campos():
    colunas = df.columns.tolist()
    estado_antigo = carregar_estado_checkboxes()
    for widget in scrollable_fields.winfo_children(): widget.destroy()
    for widget in frame_desc.winfo_children(): widget.destroy()
    entradas.clear(); check_vars.clear()

    metade = (len(colunas)+1)//2
    idx = 0
    for col in colunas:
        if col.lower()=='descrição': continue
        line = idx if idx<metade else idx-metade
        coln = 0 if idx<metade else 2
        idx+=1
        var = tk.BooleanVar(value=estado_antigo.get(col, True))
        chk = tk.Checkbutton(scrollable_fields, variable=var, command=lambda c=col: toggle_campo(c))
        chk.grid(row=line, column=coln*2, padx=5, pady=2, sticky='w')
        tk.Label(scrollable_fields, text=col).grid(row=line, column=coln*2+1, sticky='e', padx=5, pady=2)
        ent = tk.Entry(scrollable_fields, width=30)
        ent.grid(row=line, column=coln*2+2, padx=5, pady=2, sticky='we')
        scrollable_fields.grid_columnconfigure(coln*2+2, weight=1)
        entradas[col]=ent
        check_vars[col]=var
        toggle_campo(col)

    if 'Descrição' in colunas:
        col = 'Descrição'
        var = tk.BooleanVar(value=estado_antigo.get(col, True))
        chk = tk.Checkbutton(frame_desc, variable=var, command=lambda c=col: toggle_campo(c))
        chk.pack(anchor='nw', padx=5, pady=2)
        tk.Label(frame_desc, text=col).pack(anchor='nw', padx=5)
        
        # Ajuste a altura do ScrolledText
        txt = scrolledtext.ScrolledText(frame_desc, wrap='word', height=10)  # Reduza o valor de 'height'
        txt.pack(fill='both', expand=True, padx=5, pady=2)
        entradas[col] = txt
        check_vars[col] = var
        toggle_campo(col)

def atualizar_tabela(selecionar_ultimo=False, selecionar_indice=None):
    global arvore, pagina_atual
    for w in frame_tabela.winfo_children():
        w.destroy()
    # Use apenas as colunas do DataFrame
    cols = df.columns.tolist()
    scr_y = tk.Scrollbar(frame_tabela, orient='vertical')
    scr_x = tk.Scrollbar(frame_tabela, orient='horizontal')
    arvore = ttk.Treeview(frame_tabela, columns=cols, show='headings', yscrollcommand=scr_y.set, xscrollcommand=scr_x.set)
    arvore.column('#0', width=0, stretch=False)

    # Configura as colunas
    for col in df.columns:
        arvore.heading(col, text=col, command=lambda c=col: ordenar_por_coluna(c))
        arvore.column(col, width=120, anchor='center')

    scr_y.config(command=arvore.yview)
    scr_x.config(command=arvore.xview)
    scr_y.pack(side='right', fill='y')
    scr_x.pack(side='bottom', fill='x')
    arvore.pack(fill='both', expand=True)

    # Preenche a tabela com os dados
    inicio = pagina_atual * linhas_por_pagina
    fim = inicio + linhas_por_pagina
    rows = list(df.iterrows())[inicio:fim]
    for i, (index, row) in enumerate(rows):
        tag = 'oddrow' if i % 2 else 'evenrow'
        arvore.insert('', tk.END, values=list(row), tags=(tag,))

    # Vincular evento de clique para atualizar o formulário
    arvore.bind('<ButtonRelease-1>', preencher_campos)

    # Botões de navegação
    botao_anterior = tk.Button(frame_tabela, text="Anterior", command=lambda: mudar_pagina(-1))
    botao_anterior.pack(side='left')
    botao_proximo = tk.Button(frame_tabela, text="Próximo", command=lambda: mudar_pagina(1))
    botao_proximo.pack(side='right')

def mudar_pagina(direcao):
    global pagina_atual
    pagina_atual += direcao
    if pagina_atual < 0:
        pagina_atual = 0
    elif pagina_atual * linhas_por_pagina >= len(df):
        pagina_atual -= 1
    atualizar_tabela()

def alternar_estado_coluna(coluna):
    global df
    selecionados = arvore.selection()  # Obtém os itens selecionados na tabela

    if coluna not in df.columns:
        messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no DataFrame.")
        return

    if selecionados:
        # Alternar apenas a célula selecionada
        for sel in selecionados:
            idx = arvore.index(sel)  # Obtém o índice da linha selecionada
            valor_atual = df.at[idx, coluna]
            novo_valor = 'NÃO' if str(valor_atual).strip().upper() == 'SIM' else 'SIM'
            df.at[idx, coluna] = novo_valor
    else:
        # Definir todas as células da coluna para "SIM"
        df[coluna] = 'SIM'

    atualizar_tabela()  # Atualiza a tabela exibida

def definir_valores_padrao():
    global df
    valores_padrao = {
        'Peso (kg)': 0.003,
        'Comprimento (cm)': 10.00,
        'Largura (cm)': 7.00,
        'Altura (cm)': 0.01
    }
    for coluna, valor in valores_padrao.items():
        if coluna in df.columns:
            df[coluna] = valor  # Define o valor padrão para todas as linhas da coluna
        else:
            messagebox.showwarning("Aviso", f"A coluna '{coluna}' não existe no DataFrame.")
    atualizar_tabela()  # Atualiza a tabela exibida

def abrir_pesquisa_avancada():
    janela_pesquisa = tk.Toplevel(janela)
    janela_pesquisa.title("Pesquisa Avançada")
    janela_pesquisa.geometry("400x400")



    criterios = {}

    def aplicar_pesquisa():
        global df
        filtro = df.copy()
        for coluna, entrada in criterios.items():
            valor = entrada.get().strip()
            if valor:
                if coluna in df.columns:
                    filtro = filtro[filtro[coluna].astype(str).str.contains(valor, case=False, na=False)]
                else:
                    messagebox.showwarning("Aviso", f"A coluna '{coluna}' não existe no DataFrame.")
        atualizar_tabela_com_filtro(filtro)
        janela_pesquisa.destroy()

    tk.Label(janela_pesquisa, text="Preencha os critérios de pesquisa:").pack(pady=10)

    for i, coluna in enumerate(df.columns):
        tk.Label(janela_pesquisa, text=coluna).pack(anchor='w', padx=10)
        entrada = tk.Entry(janela_pesquisa, width=30)
        entrada.pack(padx=10, pady=5)
        criterios[coluna] = entrada

    tk.Button(janela_pesquisa, text="Aplicar Filtro", command=aplicar_pesquisa).pack(pady=20)

# Campos que sempre devem ser desmarcados na duplicata
CAMPOS_SEMPRE_DESMARCADOS = [
    "Nome", "Categorias", "Tags", "Título para SEO", "Descrição para SEO",
    "Marca", "Código de barras", "MPN (Cód. Exclusivo, Modelo Fabricante)",
    "Sexo", "Faixa etária", "Custo"
]
def setar_estoque_para_um():
    global df
    if 'Estoque' not in df.columns:
        messagebox.showerror("Erro", "Coluna 'Estoque' não encontrada.")
        return

    if alterar_todos_var.get():
        # Altera todos os produtos
        df.loc[df['Estoque'] != 0, 'Estoque'] = 1
    else:
        # Altera apenas os selecionados
        selecionados = arvore.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")
            return
        cols = df.columns.tolist()
        for sel in selecionados:
            idx = arvore.index(sel)
            if df.at[idx, 'Estoque'] != 0:
                df.at[idx, 'Estoque'] = 1

    atualizar_tabela()
    messagebox.showinfo("Sucesso", "Estoques diferentes de 0 foram definidos como 1.")

def duplicar_produto():
    global df
    sel = arvore.focus()
    if not sel:
        messagebox.showwarning("Aviso", "Nenhum produto selecionado para duplicar.")
        return

    valores = arvore.item(sel, 'values')
    cols = df.columns.tolist()
    if 'SKU' not in cols or 'Valor da variação 1' not in cols:
        messagebox.showerror("Erro", "Coluna 'SKU' ou 'Valor da variação 1' não encontrada.")
        return

    sku_idx = cols.index('SKU')
    variacao_idx = cols.index('Valor da variação 1')
    sku_val = valores[sku_idx]
    variacao_val = valores[variacao_idx]

    # Extrai a base da SKU (tudo antes do número sequencial)
    match = re.match(r'([A-Za-z0-9\-]+?)-(\d+)$', sku_val)
    if not match:
        messagebox.showerror("Erro", "Formato de SKU inválido.")
        return
    sku_base = match.group(1)
    numero_atual = int(match.group(2))

    # Define a variação oposta
    if str(variacao_val).strip().lower() == "dourado":
        nova_variacao = "Prata"
    elif str(variacao_val).strip().lower() == "prata":
        nova_variacao = "Dourado"
    else:
        messagebox.showerror("Erro", "Variação não reconhecida.")
        return

    # Verifica se já existe a combinação SKU-base + nova_variacao
    existe = df[
        df['SKU'].str.startswith(sku_base + "-") &
        (df['Valor da variação 1'].str.strip().str.lower() == nova_variacao.lower())
    ]
    if not existe.empty:
        messagebox.showinfo("Info", f"A variação '{nova_variacao}' já existe para esta SKU base.")
        return

    # Cria a nova linha duplicada
    idxs = df.index[df['SKU'] == sku_val].tolist()
    if not idxs:
        messagebox.showerror("Erro", "SKU não encontrada no DataFrame.")
        return
    idx = idxs[0]
    nova_linha = df.iloc[idx].copy()

    # Atualiza a variação e incrementa o número da SKU
    nova_linha['Valor da variação 1'] = nova_variacao
    nova_linha['SKU'] = f"{sku_base}-{numero_atual + 1:03}"

    # Campos que devem ser definidos como NaN
    CAMPOS_NAN = ["Nome", "Categorias", "Tags", "Título para SEO", "Descrição para SEO"]
    for col in CAMPOS_NAN:
        if col in nova_linha:
            nova_linha[col] = pd.NA

    # Adiciona a nova linha ao DataFrame
    df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
    atualizar_tabela(selecionar_ultimo=True)
def excluir_produto_por_sku():
    def confirmar_exclusao():
        sku = entry_sku.get().strip()
        if not sku:
            messagebox.showwarning("Aviso", "Digite a SKU para excluir.")
            return
        global df
        if 'SKU' not in df.columns:
            messagebox.showerror("Erro", "Coluna 'SKU' não encontrada.")
            janela_excluir.destroy()
            return
        mask = df['SKU'].astype(str) == sku
        if mask.any():
            df.drop(df[mask].index, inplace=True)
            df.reset_index(drop=True, inplace=True)
            atualizar_tabela()
            messagebox.showinfo("Sucesso", f"Produto com SKU {sku} excluído.")
        else:
            messagebox.showwarning("Aviso", f"SKU {sku} não encontrada.")
        janela_excluir.destroy()

    janela_excluir = tk.Toplevel(janela)
    janela_excluir.title("Excluir Produto por SKU")
    tk.Label(janela_excluir, text="Digite a SKU do produto a excluir:").pack(padx=10, pady=10)
    entry_sku = tk.Entry(janela_excluir)
    entry_sku.pack(padx=10, pady=5)
    tk.Button(janela_excluir, text="Excluir", command=confirmar_exclusao).pack(pady=10)

# Montagem GUI
janela=tk.Tk()
janela.title('Cadastro de Produtos CSV')
janela.state('zoomed')

filter_var = tk.StringVar()
coluna_filtro_var = tk.StringVar()

frame_filtro = tk.Frame(janela)
frame_filtro.pack(fill='x', padx=5, pady=5)

tk.Label(frame_filtro, text="Filtrar:").pack(side='left')
entry_filtro = tk.Entry(frame_filtro, textvariable=filter_var)
entry_filtro.pack(side='left', fill='x', expand=True)
filter_var.trace_add('write', lambda *args: aplicar_filtro())

tk.Label(frame_filtro, text="Coluna:").pack(side='left', padx=5)
coluna_filtro = ttk.Combobox(frame_filtro, textvariable=coluna_filtro_var, state='readonly', width=20)
coluna_filtro['values'] = []  # Inicializa vazio
coluna_filtro.pack(side='left', padx=5)

criar_menu_contexto()

paned=tk.PanedWindow(janela,orient='horizontal',sashwidth=8,sashrelief='raised')
paned.pack(fill='both',expand=False,padx=5,pady=5)

left=tk.Frame(paned)
paned.add(left)

canvas_fields=tk.Canvas(left)
scroll_y=tk.Scrollbar(left,orient='vertical',command=canvas_fields.yview)
scrollable_fields=tk.Frame(canvas_fields)
scrollable_fields.bind('<Configure>',lambda e:canvas_fields.configure(scrollregion=canvas_fields.bbox('all')))
canvas_fields.create_window((0,0),window=scrollable_fields,anchor='nw')
canvas_fields.configure(yscrollcommand=scroll_y.set)
canvas_fields.pack(side='left',fill='both',expand=True)
scroll_y.pack(side='right',fill='y')

frame_desc = tk.Frame(paned, height=200, width=400)  # Define a largura inicial do frame
paned.add(frame_desc)

# Ajusta o peso do PanedWindow para equilibrar os frames
paned.paneconfig(left, stretch="always")  # O frame esquerdo pode expandir
paned.paneconfig(frame_desc, stretch="never")  # O frame de descrição não expande

frame_tabela=tk.Frame(janela)
frame_tabela.pack(fill='both',expand=True,padx=5)

rodape = tk.Frame(janela)
rodape.pack(fill='x', pady=5)

alterar_todos_var = tk.BooleanVar(value=False)  # Variável para controlar o estado do checkbox

tk.Checkbutton(rodape, text="Alterar Todos", variable=alterar_todos_var).pack(side='left', padx=5)

# Botões no rodapé
for text, cmd in [
    ('Adicionar Produto', adicionar_linha),
    ('Alterar Produto', alterar_linha),
    ('Limpar Campos', lambda: limpar_campos(preservar_estado=True))
]:
    tk.Button(rodape, text=text, command=cmd).pack(side='left', padx=5)

tk.Button(rodape, text='Salvar CSV', command=salvar_csv, bg='#4CAF50', fg='white').pack(side='right', padx=5)
tk.Button(rodape, text='Adicionar CSV', command=carregar_csv).pack(side='right', padx=5)
tk.Button(rodape, text='Exibir na Loja: Alternar', command=lambda: alternar_estado_coluna('Exibir na loja')).pack(side='left', padx=5)
tk.Button(rodape, text='Produto Físico: Alternar', command=lambda: alternar_estado_coluna('Produto Físico')).pack(side='left', padx=5)
tk.Button(rodape, text='Definir Valores Padrão', command=definir_valores_padrao).pack(side='left', padx=5)
tk.Button(rodape, text='Duplicar Produto', command=duplicar_produto).pack(side='left', padx=5)
tk.Button(rodape, text='Excluir Produto', command=excluir_produto_por_sku, bg='#f44336', fg='white').pack(side='left', padx=5)
tk.Button(rodape, text='Setar para 1 diferente de 0 ', command=setar_estoque_para_um).pack(side='left', padx=5)
# Inicializar
criar_campos()
atualizar_tabela()
janela.bind('<Control-n>', lambda e: limpar_campos())
janela.bind('<Control-s>', lambda e: salvar_csv())
janela.bind('<Control-f>', focar_filtro)
arvore.bind('<Button-1>', alternar_checkbox)
janela.mainloop()
