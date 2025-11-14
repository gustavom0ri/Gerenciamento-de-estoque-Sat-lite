import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
from tkinter.ttk import Combobox
from PIL import Image, ImageTk
import datetime
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import shutil
import copy
import time
from collections import defaultdict

# --------------------------  CONFIGURAÇÕES GLOBAIS  --------------------------
estoque = []
removidos = []
item_selecionado = None
historico = []
painel_aberto = False
painel_minimizado = False
undo_stack = []
redo_stack = []
ultima_data_registrada = None
updating = False
last_successful_auth = 0
categorias = ["Sem Categoria"]
categoria_aberta = {}

PASTA_IMAGENS = "imagens_estoque"
CAMINHO_HISTORICO = "historico.txt"
CAMINHO_DB = "estoque.xlsx"
CAMINHO_TEMP = "estoque_temp.xlsx"
TAMANHO_CARD = 250

if not os.path.exists(PASTA_IMAGENS):
    os.makedirs(PASTA_IMAGENS)

# --------------------------  JANELA PRINCIPAL  --------------------------
root = tk.Tk()
root.title("Sistema de Estoque Satelite")
root.state("zoomed")
root.update()

# --------------------------  PLACEHOLDER DA BUSCA  --------------------------
placeholder_text = "Pesquisar itens..."
search_var = tk.StringVar()
search_active = False

# --------------------------  FUNÇÕES AUXILIARES  --------------------------
def create_padded_photoimage(image_path, size, bg_color=(44, 44, 44)):
    if not os.path.exists(image_path):
        return None
    try:
        img = Image.open(image_path)
        img.thumbnail(size)
        new_img = Image.new("RGB", size, bg_color)
        offset = ((size[0] - img.width) // 2, (size[1] - img.height) // 2)
        new_img.paste(img, offset)
        return ImageTk.PhotoImage(new_img)
    except Exception as e:
        print(f"Erro ao criar imagem padded: {e}")
        return None


def salvar_no_excel():
    try:
        with pd.ExcelWriter(CAMINHO_TEMP, engine='openpyxl') as writer:
            if estoque:
                dados_estoque = []
                for item in estoque:
                    dados_estoque.append({
                        "Nome": item["nome"],
                        "Id": item.get("id", f"ID_{len(dados_estoque) + 1}"),
                        "Quantidade": item["quantidade"],
                        "Tipo_Unidade": item["tipo_unidade"],   # NOVA COLUNA
                        "Preco": item.get("preco", None),
                        "Categoria": item.get("categoria", "Sem Categoria"),
                        "Data_Criacao": item.get("data_criacao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")),
                        "Data_Alteracao": item.get("data_alteracao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
                    })
                df_estoque = pd.DataFrame(dados_estoque)
                df_estoque.to_excel(writer, sheet_name='Estoque', index=False)

            if removidos:
                dados_removidos = []
                for item in removidos:
                    dados_removidos.append({
                        "Nome": item["nome"],
                        "Id": item.get("id", ""),
                        "Quantidade": item["quantidade"],
                        "Tipo_Unidade": item.get("tipo_unidade", "Unidade"),
                        "Preco": item.get("preco", None),
                        "Categoria": item.get("categoria", "Sem Categoria"),
                        "Data_Criacao": item.get("data_criacao", ""),
                        "Data_Alteracao": item.get("data_alteracao", ""),
                        "Data_Remocao": item.get("data_remocao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
                    })
                df_removidos = pd.DataFrame(dados_removidos)
                df_removidos.to_excel(writer, sheet_name='Removidos', index=False)

            df_categorias = pd.DataFrame({'Categoria': categorias})
            df_categorias.to_excel(writer, sheet_name='Categorias', index=False)

        if os.path.exists(CAMINHO_DB):
            os.remove(CAMINHO_DB)
        os.rename(CAMINHO_TEMP, CAMINHO_DB)
        print(f"Excel salvo com sucesso: {CAMINHO_DB}")
    except Exception as e:
        if os.path.exists(CAMINHO_TEMP):
            try: os.remove(CAMINHO_TEMP)
            except: pass
        messagebox.showerror("Erro", f"Não foi possível salvar o Excel:\n{e}\nFeche o arquivo manualmente.")


def carregar_do_excel():
    global estoque, removidos, categorias
    estoque.clear()
    removidos.clear()

    if not os.path.exists(CAMINHO_DB):
        print(f"Arquivo {CAMINHO_DB} não encontrado. Iniciando com estoque vazio.")
        categorias = ["Sem Categoria"]
        return

    try:
        with pd.ExcelFile(CAMINHO_DB) as xls:
            if 'Estoque' in xls.sheet_names:
                df_estoque = pd.read_excel(xls, 'Estoque')
                seen_ids = set()
                for _, row in df_estoque.iterrows():
                    item_id = str(row.get("Id", f"ID_{len(estoque) + 1}"))
                    if item_id in seen_ids: continue
                    seen_ids.add(item_id)

                    nome_arquivo = f"{row['Nome']}_{item_id}.jpg"
                    caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)
                    image_path = caminho_imagem if os.path.exists(caminho_imagem) else None

                    preco = row.get("Preco", None)
                    if pd.isna(preco): preco = None
                    else: preco = float(preco)

                    # NOVO: Tipo de unidade
                    tipo_unidade = row.get("Tipo_Unidade", "Unidade")
                    if pd.isna(tipo_unidade): tipo_unidade = "Unidade"

                    item = {
                        "image_path": image_path,
                        "nome": str(row["Nome"]),
                        "quantidade": float(row["Quantidade"]),
                        "tipo_unidade": tipo_unidade,        # <<< NOVO
                        "preco": preco,
                        "categoria": str(row.get("Categoria", "Sem Categoria")),
                        "var_esq": tk.IntVar(value=1),
                        "var_dir": tk.IntVar(value=1),
                        "id": item_id,
                        "data_criacao": str(row.get("Data_Criacao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))),
                        "data_alteracao": str(row.get("Data_Alteracao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))
                    }
                    estoque.append(item)

            if 'Removidos' in xls.sheet_names:
                df_removidos = pd.read_excel(xls, 'Removidos')
                for _, row in df_removidos.iterrows():
                    preco = row.get("Preco", None)
                    if pd.isna(preco): preco = None
                    else: preco = float(preco)

                    item = {
                        "nome": str(row["Nome"]),
                        "quantidade": float(row["Quantidade"]),
                        "tipo_unidade": row.get("Tipo_Unidade", "Unidade"),
                        "preco": preco,
                        "categoria": str(row.get("Categoria", "Sem Categoria")),
                        "id": str(row.get("Id", "")),
                        "data_criacao": str(row.get("Data_Criacao", "")),
                        "data_alteracao": str(row.get("Data_Alteracao", "")),
                        "data_remocao": str(row.get("Data_Remocao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))
                    }
                    removidos.append(item)

            if 'Categorias' in xls.sheet_names:
                df_categorias = pd.read_excel(xls, 'Categorias')
                categorias = list(df_categorias['Categoria'].unique())
                if "Sem Categoria" not in categorias:
                    categorias.insert(0, "Sem Categoria")
            else:
                categorias = ["Sem Categoria"]
    except Exception as e:
        print(f"Erro ao carregar Excel: {e}")
        messagebox.showerror("Erro", f"Falha ao carregar o arquivo de estoque: {e}")
        categorias = ["Sem Categoria"]


def salvar_imagem(caminho_original, nome_item, item_id):
    if not os.path.exists(caminho_original): return None
    nome_arquivo = f"{nome_item}_{item_id}.jpg"
    caminho_destino = os.path.join(PASTA_IMAGENS, nome_arquivo)
    try:
        shutil.copy2(caminho_original, caminho_destino)
        return caminho_destino
    except Exception as e:
        print(f"Erro ao salvar imagem: {e}")
        return None


def salvar_estado():
    estado = {
        'estoque': [],
        'removidos': []
    }
    for item in estoque:
        estado_estoque = {
            'nome': item['nome'],
            'quantidade': item['quantidade'],
            'preco': item.get('preco', None),
            'categoria': item.get('categoria', "Sem Categoria"),
            'id': item['id'],
            'data_criacao': item['data_criacao'],
            'data_alteracao': item['data_alteracao'],
            'image_path': item.get('image_path'),  # salva o caminho, não o objeto imagem
            # NÃO salva var_esq/var_dir (são tk.IntVar → não picklable)
        }
        estado['estoque'].append(estado_estoque)

    # Copia removidos (sem objetos Tkinter)
    for item in removidos:
        estado_removido = {
            'nome': item['nome'],
            'quantidade': item['quantidade'],
            'preco': item.get('preco', None),
            'categoria': item.get('categoria', "Sem Categoria"),
            'id': item.get('id', ''),
            'data_criacao': item.get('data_criacao', ''),
            'data_alteracao': item.get('data_alteracao', ''),
            'data_remocao': item.get('data_remocao', '')
        }
        estado['removidos'].append(estado_removido)

    undo_stack.append(estado)
    if len(undo_stack) > 20:
        undo_stack.pop(0)


def recarregar_imagens_estoque(estoque_dados):
    novo_estoque = []
    for item_data in estoque_dados:
        nome_arquivo = f"{item_data['nome']}_{item_data['id']}.jpg"
        caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)
        image_path = caminho_imagem if os.path.exists(caminho_imagem) else item_data.get('image_path')

        item = {
            "image_path": image_path,
            "nome": item_data['nome'],
            "quantidade": item_data['quantidade'],
            "preco": item_data.get('preco', None),
            "categoria": item_data.get('categoria', "Sem Categoria"),
            "var_esq": tk.IntVar(value=1),   # RECRIA os IntVar
            "var_dir": tk.IntVar(value=1),   # RECRIA os IntVar
            "id": item_data['id'],
            "data_criacao": item_data['data_criacao'],
            "data_alteracao": item_data['data_alteracao']
        }
        novo_estoque.append(item)
    return novo_estoque


def undo(event=None):
    global estoque, removidos
    if not undo_stack:
        messagebox.showinfo("Undo", "Nada para desfazer!")
        return
    estado_atual = {'estoque': [], 'removidos': copy.deepcopy(removidos)}
    for item in estoque:
        estado_estoque = {k: item.get(k) for k in ['nome', 'quantidade', 'preco', 'categoria', 'id', 'data_criacao', 'data_alteracao']}
        estado_atual['estoque'].append(estado_estoque)
    redo_stack.append(estado_atual)
    if len(redo_stack) > 20:
        redo_stack.pop(0)

    estado_anterior = undo_stack.pop()
    removidos = copy.deepcopy(estado_anterior['removidos'])
    estoque = recarregar_imagens_estoque(estado_anterior['estoque'])
    salvar_no_excel()
    registrar_historico("Desfeita a última ação")
    atualizar_tela()


def redo(event=None):
    global estoque, removidos
    if not redo_stack:
        messagebox.showinfo("Redo", "Nada para refazer!")
        return
    estado_atual = {'estoque': [], 'removidos': copy.deepcopy(removidos)}
    for item in estoque:
        estado_estoque = {k: item.get(k) for k in ['nome', 'quantidade', 'preco', 'categoria', 'id', 'data_criacao', 'data_alteracao']}
        estado_atual['estoque'].append(estado_estoque)
    undo_stack.append(estado_atual)
    if len(undo_stack) > 20:
        undo_stack.pop(0)

    estado_redo = redo_stack.pop()
    removidos = copy.deepcopy(estado_redo['removidos'])
    estoque = recarregar_imagens_estoque(estado_redo['estoque'])
    salvar_no_excel()
    registrar_historico("Refeita a ação desfeita")
    atualizar_tela()


def exportar_historico():
    if not historico:
        messagebox.showwarning("Exportar", "Histórico vazio!")
        return
    caminho = filedialog.asksaveasfilename(title="Salvar Histórico", defaultextension=".txt",
                                           filetypes=[("Texto", "*.txt")])
    if not caminho:
        return
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            for acao in historico:
                f.write(acao + "\n")
        messagebox.showinfo("Exportar", f"Histórico exportado em:\n{caminho}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar: {e}")


def registrar_historico(acao):
    global ultima_data_registrada
    agora = datetime.datetime.now()
    data_atual = agora.strftime("%d/%m/%Y")
    hora = agora.strftime("%H:%M:%S")
    acao_completa = f"[{hora}] {acao}"

    if ultima_data_registrada and ultima_data_registrada != data_atual:
        historico.append("--------")
        try:
            with open(CAMINHO_HISTORICO, "a", encoding="utf-8") as f:
                f.write("--------\n")
        except:
            pass

    historico.append(acao_completa)
    ultima_data_registrada = data_atual
    try:
        with open(CAMINHO_HISTORICO, "a", encoding="utf-8") as f:
            f.write(acao_completa + "\n")
    except:
        pass

    if painel_aberto:
        atualizar_historico()


# --------------------------  PAINEL HISTÓRICO  --------------------------
def toggle_historico():
    global painel_aberto, painel_minimizado
    if painel_minimizado:
        expandir_painel()
    else:
        if painel_aberto:
            painel_historico.pack_forget()
            painel_aberto = False
            btn_toggle_historico.config(text="Histórico")
        else:
            painel_historico.pack(side="right", fill="y", padx=(0, 0))
            painel_aberto = True
            btn_toggle_historico.config(text="Fechar")
            atualizar_historico()
    ajustar_modo_visualizacao()


def minimizar_painel():
    global painel_minimizado, painel_aberto
    if not painel_aberto:
        return
    painel_historico.pack_forget()
    painel_minimizado = True
    btn_toggle_historico.config(text="Histórico")
    ajustar_modo_visualizacao()


def expandir_painel():
    global painel_minimizado, painel_aberto
    if painel_minimizado:
        painel_historico.pack(side="right", fill="y", padx=(0, 0))
        painel_minimizado = False
        painel_aberto = True
        btn_toggle_historico.config(text="Fechar")
        atualizar_historico()
        ajustar_modo_visualizacao()


def atualizar_historico():
    for widget in painel_historico.winfo_children():
        widget.destroy()

    tk.Label(painel_historico, text="Histórico de Ações", font=("Arial", 14, "bold"),
             bg="#222", fg="#00d4ff").pack(pady=10)

    frame_principal = tk.Frame(painel_historico, bg="#222")
    frame_principal.pack(fill="both", expand=True, padx=10)

    canvas_hist = tk.Canvas(frame_principal, bg="#222", highlightthickness=0, width=280)
    scrollbar = tk.Scrollbar(frame_principal, orient="vertical", command=canvas_hist.yview)
    scroll_frame_hist = tk.Frame(canvas_hist, bg="#222")

    scroll_frame_hist.bind(
        "<Configure>",
        lambda e: canvas_hist.configure(scrollregion=canvas_hist.bbox("all"))
    )
    canvas_hist.create_window((0, 0), window=scroll_frame_hist, anchor="nw")
    canvas_hist.configure(yscrollcommand=scrollbar.set)

    canvas_hist.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel_hist(event):
        canvas_hist.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas_hist.bind("<MouseWheel>", _on_mousewheel_hist)

    for acao in historico[::-1]:
        bg = "#333" if acao != "--------" else "#222"
        fg = "white" if acao != "--------" else "#888"
        tk.Label(scroll_frame_hist, text=acao, anchor="w", justify="left",
                 wraplength=250, font=("Arial", 10), bg=bg, fg=fg).pack(fill="x", padx=2, pady=1)

    tk.Button(painel_historico, text="Exportar Log", command=exportar_historico,
              bg="#17a2b8", fg="white", font=("Arial", 11, "bold"), relief="flat").pack(side="bottom", pady=8, fill="x")


# --------------------------  BUSCA (PLACEHOLDER)  --------------------------
def on_search_focus_in(event):
    global search_active
    if not search_active:
        entry_search.delete(0, tk.END)
        entry_search.config(fg="white")
        search_active = True


def on_search_focus_out(event):
    global search_active
    if not entry_search.get().strip():
        entry_search.insert(0, placeholder_text)
        entry_search.config(fg="#aaa")
        search_active = False
        if painel_aberto and not painel_minimizado:
            minimizar_painel()


def on_search_keyrelease(event):
    termo = search_var.get().strip()
    if termo and termo != placeholder_text:
        if painel_minimizado:
            expandir_painel()
    atualizar_tela()


# --------------------------  AUTENTICAÇÃO  --------------------------
def open_password_form(action_title):
    global last_successful_auth
    if time.time() - last_successful_auth < 60:
        return True

    topo = tk.Toplevel(root)
    topo.title("Autenticação")
    topo.geometry("400x250")
    topo.configure(bg="#2c2c2c")
    x = (root.winfo_screenwidth() - 400) // 2
    y = (root.winfo_screenheight() - 250) // 2
    topo.geometry(f"+{x}+{y}")
    topo.transient(root)
    topo.grab_set()

    tk.Label(topo, text=f"Senha para {action_title}", font=("Arial", 16, "bold"),
             bg="#2c2c2c", fg="#00d4ff").pack(pady=20)

    entry = tk.Entry(topo, show="*", font=("Arial", 12), bg="#444", fg="white",
                     insertbackground="white")
    entry.pack(pady=10, fill="x", padx=40)
    entry.focus()

    result = [False]

    def confirmar():
        if entry.get() == "senhaadmsatelite":
            global last_successful_auth
            last_successful_auth = time.time()
            result[0] = True
            topo.destroy()
        else:
            messagebox.showerror("Erro", "Senha incorreta!")
            entry.delete(0, tk.END)

    tk.Button(topo, text="Confirmar", command=confirmar, bg="#28a745",
              fg="white").pack(side="left", padx=20, pady=20)
    tk.Button(topo, text="Cancelar", command=topo.destroy, bg="#dc3545",
              fg="white").pack(side="right", padx=20, pady=20)

    topo.bind("<Return>", lambda e: confirmar())
    root.wait_window(topo)
    return result[0]

# ======================  FUNÇÃO PARA VERIFICAR ID DUPLICADO  ======================

def verificar_id_duplicado(id_proposto, item_atual=None):
    """Retorna o nome do item que já tem esse ID, ou None se disponível"""
    for item in estoque:
        if item["id"] == id_proposto:
            if item_atual is not None and item is item_atual:
                continue
            return item["nome"]
    return None

# ======================  FUNÇÃO PARA VERIFICAR ID DUPLICADO  ======================
def verificar_id_duplicado(id_proposto, item_atual=None):
    """Retorna None se OK, ou o nome do item que já tem esse ID"""
    for item in estoque:
        if item["id"] == id_proposto:
            if item_atual is not None and item is item_atual:
                continue  # permite manter o mesmo ID ao editar o próprio item
            return item["nome"]
    return None


def open_item_form(item=None):
    global categorias
    is_edit = item is not None
    title = "Editar Item" if is_edit else "Adicionar Item"

    topo = tk.Toplevel(root)
    topo.title(title)
    topo.configure(bg="#2c2c2c")
    topo.transient(root)
    topo.grab_set()
    topo.geometry("1100x620")
    topo.update_idletasks()
    x = (root.winfo_screenwidth() - 1100) // 2
    y = (root.winfo_screenheight() - 620) // 2
    topo.geometry(f"+{x}+{y}")

    # ---------- LADO ESQUERDO ----------
    frame_esq = tk.Frame(topo, bg="#2c2c2c")
    frame_esq.pack(side="left", fill="both", expand=True, padx=20, pady=20)

    # ---------- LADO DIREITO ----------
    frame_dir = tk.Frame(topo, bg="#333", bd=2, relief="groove")
    frame_dir.pack(side="right", fill="both", expand=True, padx=20, pady=20)

    tk.Label(frame_esq, text=title, font=("Arial", 20, "bold"), bg="#2c2c2c", fg="#00d4ff").pack(pady=(0, 20))

    # ---------- ID COM VALIDAÇÃO OBRIGATÓRIA ----------
    tk.Label(frame_esq, text="ID do Item (ex: ID_1):", bg="#2c2c2c", fg="white", font=("Arial", 12)).pack(anchor="w")
    entry_id = tk.Entry(frame_esq, font=("Arial", 12), bg="#444", fg="white", width=20)
    entry_id.pack(anchor="w", pady=5)
    lbl_aviso_id = tk.Label(frame_esq, text="", bg="#2c2c2c", fg="#ff6b6b", font=("Arial", 10))
    lbl_aviso_id.pack(anchor="w")

    if is_edit:
        entry_id.insert(0, item["id"])
    else:
        proximo = 1
        while any(i["id"] == f"ID_{proximo}" for i in estoque):
            proximo += 1
        entry_id.insert(0, f"ID_{proximo}")

    def validar_id(event=None):
        texto = entry_id.get().strip()
        if not texto:
            lbl_aviso_id.config(text="ID obrigatório!", fg="#ff6b6b")
            return

        if not texto.startswith("ID_"):
            lbl_aviso_id.config(text="ID deve começar com 'ID_'", fg="#ff6b6b")
            return

        try:
            numero = int(texto[3:])
            if numero <= 0:
                raise ValueError
        except:
            lbl_aviso_id.config(text="Após 'ID_' deve ser um número positivo", fg="#ff6b6b")
            return

        duplicado = verificar_id_duplicado(texto, item)
        if duplicado:
            lbl_aviso_id.config(text=f"ID já usado por: {duplicado}", fg="#ff6b6b")
        else:
            lbl_aviso_id.config(text="ID válido e disponível", fg="#66ff66")

    entry_id.bind("<KeyRelease>", validar_id)
    validar_id()  # Validação inicial

    # ---------- Nome ----------
    tk.Label(frame_esq, text="Nome:", bg="#2c2c2c", fg="white", font=("Arial", 12)).pack(anchor="w", pady=(15,0))
    entry_nome = tk.Entry(frame_esq, font=("Arial", 12), bg="#444", fg="white")
    entry_nome.pack(fill="x", pady=5)
    if is_edit: entry_nome.insert(0, item["nome"])

    # ---------- Quantidade + Tipo ----------
    tk.Label(frame_esq, text="Quantidade:", bg="#2c2c2c", fg="white", font=("Arial", 12)).pack(anchor="w", pady=(15,0))
    frame_qtd = tk.Frame(frame_esq, bg="#2c2c2c")
    frame_qtd.pack(fill="x", pady=5)
    entry_quantidade = tk.Entry(frame_qtd, font=("Arial", 12), bg="#444", fg="white", width=15)
    entry_quantidade.pack(side="left")
    var_tipo = tk.StringVar(value=item["tipo_unidade"] if is_edit else "Unidade")
    tk.Radiobutton(frame_qtd, text="Unidade", variable=var_tipo, value="Unidade", bg="#2c2c2c", fg="white", selectcolor="#444").pack(side="left", padx=10)
    tk.Radiobutton(frame_qtd, text="Kg", variable=var_tipo, value="Kg", bg="#2c2c2c", fg="white", selectcolor="#444").pack(side="left")
    if is_edit: entry_quantidade.insert(0, str(item["quantidade"]))

    # ---------- Preço ----------
    tk.Label(frame_esq, text="Preço (opcional):", bg="#2c2c2c", fg="white", font=("Arial", 12)).pack(anchor="w", pady=(15,0))
    entry_preco = tk.Entry(frame_esq, font=("Arial", 12), bg="#444", fg="white")
    entry_preco.pack(fill="x", pady=5)
    if is_edit and item.get("preco") is not None:
        entry_preco.insert(0, str(item["preco"]))

    # ---------- Categoria ----------
    tk.Label(frame_esq, text="Categoria:", bg="#2c2c2c", fg="white", font=("Arial", 12)).pack(anchor="w", pady=(15,0))
    frame_categoria = tk.Frame(frame_esq, bg="#2c2c2c")
    frame_categoria.pack(fill="x", pady=5)
    combobox_categoria = Combobox(frame_categoria, values=categorias, state="readonly", font=("Arial", 11))
    combobox_categoria.pack(side="left", fill="x", expand=True)
    combobox_categoria.set(item.get("categoria", "Sem Categoria") if is_edit else "Sem Categoria")

    def add_nova_categoria():
        nova = simpledialog.askstring("Nova Categoria", "Nome:", parent=topo)
        if nova and nova.strip() and nova.strip() not in categorias:
            categorias.append(nova.strip())
            combobox_categoria['values'] = categorias
            combobox_categoria.set(nova.strip())
            salvar_no_excel()
            registrar_historico(f"Adicionada categoria '{nova.strip()}'")

    tk.Button(frame_categoria, text="+", command=add_nova_categoria,
              bg="#17a2b8", fg="white", font=("Arial", 12, "bold"), width=3).pack(side="right", padx=5)

    # ---------- IMAGEM ----------
    IMG_SIZE = 400
    canvas_imagem = tk.Canvas(frame_dir, width=IMG_SIZE, height=IMG_SIZE, bg="#444", highlightthickness=0)
    canvas_imagem.pack(expand=True, padx=20, pady=20)

    image_path_var = tk.StringVar(value=item.get("image_path", "") if is_edit else "")
    current_photo = None

    def update_image_preview():
        nonlocal current_photo
        caminho = image_path_var.get()
        canvas_imagem.delete("all")
        if caminho and os.path.exists(caminho):
            try:
                img = Image.open(caminho)
                img.thumbnail((IMG_SIZE-40, IMG_SIZE-40), Image.Resampling.LANCZOS)
                bg_img = Image.new("RGB", (IMG_SIZE, IMG_SIZE), "#444")
                offset = ((IMG_SIZE - img.width)//2, (IMG_SIZE - img.height)//2)
                bg_img.paste(img, offset)
                current_photo = ImageTk.PhotoImage(bg_img)
                canvas_imagem.create_image(IMG_SIZE//2, IMG_SIZE//2, image=current_photo)
            except:
                canvas_imagem.create_text(IMG_SIZE//2, IMG_SIZE//2, text="Erro na imagem", fill="red", font=("Arial", 14))
        else:
            canvas_imagem.create_text(IMG_SIZE//2, IMG_SIZE//2, text="Nenhuma imagem\nClique abaixo para selecionar", fill="#aaa", font=("Arial", 14))

    if is_edit and item.get("image_path"):
        image_path_var.set(item["image_path"])
    update_image_preview()

    def selecionar_imagem():
        caminho = filedialog.askopenfilename(title="Escolha a foto do item",
                                            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif")], parent=topo)
        if caminho:
            image_path_var.set(caminho)
            update_image_preview()

    tk.Button(frame_dir, text="Selecionar Imagem", command=selecionar_imagem,
              bg="#17a2b8", fg="white", font=("Arial", 14, "bold"), padx=30, pady=12).pack(side="bottom", pady=20)

    # ---------- BOTÕES: SALVAR E CANCELAR (VERTICAL, CENTRALIZADOS) ----------
    frame_botoes = tk.Frame(topo, bg="#2c2c2c")
    frame_botoes.pack(side="bottom", fill="x", pady=20, padx=20)

    frame_botoes_interno = tk.Frame(frame_botoes, bg="#2c2c2c")
    frame_botoes_interno.pack(expand=True)

    def confirmar():
        novo_id = entry_id.get().strip()
        if not novo_id.startswith("ID_") or not novo_id[3:].isdigit() or int(novo_id[3:]) <= 0:
            messagebox.showerror("ID Inválido", "Use o formato: ID_1, ID_2, etc.")
            return

        duplicado = verificar_id_duplicado(novo_id, item)
        if duplicado:
            messagebox.showerror("ID Duplicado", f"Este ID já está em uso por:\n→ {duplicado}")
            return

        nome = entry_nome.get().strip()
        if not nome:
            messagebox.showerror("Erro", "O nome é obrigatório!")
            return

        try:
            qtd = float(entry_quantidade.get())
            if qtd < 0: raise ValueError
        except:
            messagebox.showerror("Erro", "Quantidade inválida!")
            return

        preco_str = entry_preco.get().strip()
        preco = None if not preco_str else float(preco_str or 0)
        if preco is not None and preco < 0:
            messagebox.showerror("Erro", "Preço não pode ser negativo!")
            return

        categoria = combobox_categoria.get() or "Sem Categoria"
        tipo_unidade = var_tipo.get()
        new_image_path = image_path_var.get()

        if not new_image_path and not is_edit:
            messagebox.showerror("Erro", "Selecione uma imagem!")
            return

        salvar_estado()

        if is_edit:
            nome_antigo = item["nome"]
            id_antigo = item["id"]
            image_path_antigo = item.get("image_path")

            if novo_id != id_antigo and image_path_antigo:
                antigo_nome_arq = os.path.basename(image_path_antigo)
                novo_nome_arq = f"{nome}_{novo_id}.jpg"
                novo_caminho = os.path.join(PASTA_IMAGENS, novo_nome_arq)
                if os.path.exists(image_path_antigo):
                    shutil.move(image_path_antigo, novo_caminho)
                item["image_path"] = novo_caminho

            item.update({
                "id": novo_id,
                "nome": nome,
                "quantidade": qtd,
                "tipo_unidade": tipo_unidade,
                "preco": preco,
                "categoria": categoria,
                "data_alteracao": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            })

            if new_image_path and new_image_path != image_path_antigo and new_image_path != item["image_path"]:
                if image_path_antigo and os.path.exists(image_path_antigo):
                    os.remove(image_path_antigo)
                item["image_path"] = salvar_imagem(new_image_path, nome, novo_id)

            registrar_historico(f"Editado '{nome_antigo}' → '{nome}' (ID: {id_antigo} → {novo_id})")
        else:
            saved_path = salvar_imagem(new_image_path, nome, novo_id)
            if not saved_path:
                messagebox.showerror("Erro", "Falha ao salvar imagem!")
                return

            new_item = {
                "image_path": saved_path,
                "nome": nome,
                "quantidade": qtd,
                "tipo_unidade": tipo_unidade,
                "preco": preco,
                "categoria": categoria,
                "var_esq": tk.IntVar(value=1),
                "var_dir": tk.IntVar(value=1),
                "id": novo_id,
                "data_criacao": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                "data_alteracao": datetime.datetime.now().strftime("%d/%m/%Y %H:%،M")
            }
            estoque.append(new_item)
            registrar_historico(f"Adicionado '{nome}' (ID: {novo_id})")

        salvar_no_excel()
        atualizar_tela()
        topo.destroy()

    # BOTÃO SALVAR (em cima)
    tk.Button(frame_botoes_interno, text="SALVAR", command=confirmar,
              bg="#28a745", fg="white", font=("Arial", 14, "bold"),
              padx=40, pady=15).pack(pady=(0, 8))

    # BOTÃO CANCELAR (em baixo, maior)
    tk.Button(frame_botoes_interno, text="CANCELAR", command=topo.destroy,
              bg="#dc3545", fg="white", font=("Arial", 16, "bold"),
              padx=60, pady=18).pack()

    # Enter = Salvar
    topo.bind("<Return>", lambda e: confirmar())


def adicionar_item():
    if not open_password_form("Adicionar Item"):
        return
    open_item_form()


def editar_item():
    global item_selecionado
    if not item_selecionado:
        messagebox.showwarning("Editar", "Nenhum item selecionado!")
        return
    if not open_password_form("Editar Item"):
        return
    open_item_form(item=item_selecionado)


def remover_item():
    global item_selecionado
    if not item_selecionado:
        messagebox.showwarning("Remover", "Nenhum item selecionado!")
        return
    if not open_password_form("Remover Item"):
        return

    salvar_estado()
    item_removido = {k: item_selecionado[k] for k in item_selecionado}
    item_removido["data_remocao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    removidos.append(item_removido)
    registrar_historico(f"Removido '{item_selecionado['nome']}' (ID:{item_selecionado['id']})")
    estoque.remove(item_selecionado)
    item_selecionado = None
    salvar_no_excel()
    atualizar_tela()


def restaurar_item():
    if not removidos:
        messagebox.showinfo("Restaurar", "Nenhum item removido para restaurar!")
        return

    topo = tk.Toplevel(root)
    topo.title("Restaurar Item")
    topo.geometry("500x400")
    listbox = tk.Listbox(topo, font=("Arial", 10), height=15)
    for idx, item in enumerate(removidos):
        texto = f"{item['id']} - {item['nome']} (Qtd: {item['quantidade']})"
        texto += f"\n   Removido: {item.get('data_remocao', 'N/A')}"
        listbox.insert(tk.END, texto)
    listbox.pack(fill="both", expand=True, padx=10, pady=10)

    def confirmar():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Seleção", "Selecione um item!")
            return
        idx = sel[0]
        item_restaurar = removidos.pop(idx)

        nome_arquivo = f"{item_restaurar['nome']}_{item_restaurar['id']}.jpg"
        caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)
        image_path = caminho_imagem if os.path.exists(caminho_imagem) else None

        novo_item = {
            "image_path": image_path,
            "nome": item_restaurar["nome"],
            "quantidade": item_restaurar["quantidade"],
            "preco": item_restaurar.get("preco", None),
            "categoria": item_restaurar.get("categoria", "Sem Categoria"),
            "var_esq": tk.IntVar(value=1),
            "var_dir": tk.IntVar(value=1),
            "id": item_restaurar["id"],
            "data_criacao": item_restaurar["data_criacao"],
            "data_alteracao": item_restaurar["data_alteracao"]
        }

        salvar_estado()
        estoque.append(novo_item)
        salvar_no_excel()
        registrar_historico(f"Restaurado '{novo_item['nome']}' (ID:{novo_item['id']})")
        atualizar_tela()
        topo.destroy()

    tk.Button(topo, text="Restaurar", command=confirmar, bg="#28a745",
              fg="white").pack(side="left", padx=20, pady=10)
    tk.Button(topo, text="Cancelar", command=topo.destroy, bg="#6c757d",
              fg="white", font=("Arial", 12, "bold"), ipadx=60, ipady=8).pack(side="right", padx=20, pady=10)


def adicionar_quantidade(item):
    try:
        valor = item["var_dir"].get()
        if valor <= 0:
            raise ValueError
        salvar_estado()
        item["quantidade"] += valor
        item["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        item["var_dir"].set(1)
        salvar_no_excel()
        registrar_historico(f"+{valor} em '{item['nome']}' → {item['quantidade']}")
        atualizar_tela()
    except:
        messagebox.showwarning("Aviso", "Digite um número maior que 0!")


def subtrair_quantidade(item):
    try:
        valor = item["var_esq"].get()
        if valor <= 0:
            raise ValueError
        salvar_estado()
        nova_qtd = max(0, item["quantidade"] - valor)
        item["quantidade"] = nova_qtd
        item["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        item["var_esq"].set(1)
        salvar_no_excel()
        registrar_historico(f"-{valor} de '{item['nome']}' → {nova_qtd}")
        atualizar_tela()
    except:
        messagebox.showwarning("Aviso", "Digite um número maior que 0!")


def selecionar_item(item):
    global item_selecionado
    item_selecionado = None if item_selecionado == item else item
    atualizar_tela()


def exportar_estoque():
    if not estoque:
        messagebox.showwarning("Exportar", "Nenhum item no estoque!")
        return

    df = pd.DataFrame([{
        "ID": i["id"],
        "Nome": i["nome"],
        "Quantidade": i["quantidade"],
        "Preco": i.get("preco", None),
        "Categoria": i.get("categoria", "Sem Categoria"),
        "Data_Criacao": i["data_criacao"],
        "Data_Alteracao": i["data_alteracao"]
    } for i in estoque])

    menu = tk.Toplevel(root)
    menu.title("Escolher formato")
    menu.geometry("300x200")
    x = (root.winfo_screenwidth() - 300) // 2
    y = (root.winfo_screenheight() - 200) // 2
    menu.geometry(f"+{x}+{y}")

    tk.Button(menu, text="Excel (.xlsx)", command=lambda: salvar_arquivo(df.to_excel, ".xlsx"), width=25).pack(pady=5)
    tk.Button(menu, text="CSV (.csv)", command=lambda: salvar_arquivo(df.to_csv, ".csv"), width=25).pack(pady=5)
    tk.Button(menu, text="TXT (.txt)", command=lambda: salvar_txt(df), width=25).pack(pady=5)
    tk.Button(menu, text="PDF (.pdf)", command=lambda: salvar_pdf(df), width=25).pack(pady=5)

    def salvar_arquivo(func, ext):
        caminho = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[(ext[1:], f"*{ext}")])
        if caminho:
            func(caminho, index=False)
            messagebox.showinfo("Exportar", "Exportado com sucesso!")
        menu.destroy()

    def salvar_txt(df):
        caminho = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("TXT", "*.txt")])
        if caminho:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write("ESTOQUE SATELITE\n" + "="*50 + "\n\n")
                for i in estoque:
                    f.write(f"ID: {i['id']}\nNome: {i['nome']}\nQtd: {i['quantidade']}\nPreço: {i.get('preco','N/A')}\nCat: {i.get('categoria','Sem Categoria')}\n\n")
            messagebox.showinfo("Exportar", "Exportado com sucesso!")
        menu.destroy()

    def salvar_pdf(df):
        caminho = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if caminho:
            c = canvas.Canvas(caminho, pagesize=letter)
            y = 750
            for i in estoque:
                c.drawString(50, y, f"{i['id']} - {i['nome']} | Qtd: {i['quantidade']} | Preço: {i.get('preco','N/A')}")
                y -= 20
                if y < 50:
                    c.showPage()
                    y = 750
            c.save()
            messagebox.showinfo("Exportar", "Exportado com sucesso!")
        menu.destroy()


# --------------------------  LAYOUT RESPONSIVO  --------------------------
MIN_CARD_WIDTH = 300
MAX_CARD_WIDTH = 400
CARD_PADDING = 20
pending_adjust = None
prev_width = 0


def compute_layout(largura_disponivel):
    cols = max(1, largura_disponivel // (MIN_CARD_WIDTH + CARD_PADDING))
    card_w = min(MAX_CARD_WIDTH, (largura_disponivel - cols * CARD_PADDING) // cols)
    return cols, card_w, cols == 1


def ajustar_modo_visualizacao(event=None):
    global pending_adjust, prev_width
    def do():
        global pending_adjust, prev_width
        pending_adjust = None
        larg = frame_conteudo.winfo_width() or root.winfo_width()
        if painel_aberto and not painel_minimizado:
            larg -= 300
        if abs(larg - prev_width) > 50:
            prev_width = larg
            atualizar_tela()
    if pending_adjust:
        root.after_cancel(pending_adjust)
    pending_adjust = root.after(300, do)


# --------------------------  ATUALIZAÇÃO DA TELA (CORRIGIDA!)  --------------------------
def toggle_categoria(cat):
    categoria_aberta[cat] = not categoria_aberta.get(cat, True)
    atualizar_tela()


def atualizar_tela():
    global updating, TAMANHO_CARD
    if updating: return
    updating = True

    for w in scroll_frame.winfo_children():
        w.destroy()

    termo = search_var.get().strip().lower()
    if termo == placeholder_text.lower(): termo = ""

    categorias_dict = defaultdict(list)
    for item in estoque:
        if not termo or termo in item["nome"].lower() or termo in item["id"].lower():
            categorias_dict[item.get("categoria", "Sem Categoria")].append(item)

    if not any(categorias_dict.values()):
        tk.Label(scroll_frame, text="Nenhum item encontrado", font=("Arial", 16), fg="#888", bg="#111").pack(expand=True)
        conteudo_canvas.configure(scrollregion=conteudo_canvas.bbox("all"))
        updating = False
        return

    largura = frame_conteudo.winfo_width() or root.winfo_width()
    if painel_aberto and not painel_minimizado:
        largura -= 300
    num_cols, card_w, use_list = compute_layout(largura)
    TAMANHO_CARD = card_w

    row = 0
    for cat, itens in sorted(categorias_dict.items()):
        frame_cat = tk.Frame(scroll_frame, bg="#111")
        frame_cat.grid(row=row, column=0, columnspan=num_cols, sticky="ew", pady=(15,5), padx=10)
        aberto = categoria_aberta.get(cat, True)
        icone = "▼" if aberto else "▶"
        btn_cat = tk.Button(frame_cat, text=f"{icone} {cat} ({len(itens)})", font=("Arial", 14, "bold"),
                            bg="#111", fg="#00d4ff", anchor="w", relief="flat",
                            command=lambda c=cat: toggle_categoria(c))
        btn_cat.pack(side="left", fill="x", expand=True)
        row += 1
        if not aberto: continue

        col = 0
        for item in itens:
            if use_list:
                # modo lista (não alterado)
                pass
            else:
                # MODO CARD - exibe quantidade com unidade
                frame_principal = tk.Frame(scroll_frame, width=largura//num_cols, height=TAMANHO_CARD+50, bg="#111")
                frame_principal.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                frame_principal.grid_propagate(False)

                cor = "#4444aa" if item_selecionado == item else "#2c2c2c"
                frame_item = tk.Frame(frame_principal, bg=cor, bd=3, relief="ridge")
                frame_item.pack(expand=True, fill="both", padx=5, pady=5)

                if item.get("image_path") and os.path.exists(item["image_path"]):
                    size = (TAMANHO_CARD-20, TAMANHO_CARD-100)
                    img = create_padded_photoimage(item["image_path"], size)
                    if img:
                        lbl = tk.Label(frame_item, image=img, bg=cor)
                        lbl.image = img
                        lbl.pack(pady=5)

                tk.Label(frame_item, text=f"{item['id']}\n{item['nome']}", font=("Arial", 10, "bold"), fg="#00d4ff",
                         bg=cor, justify="center", wraplength=TAMANHO_CARD-20).pack(pady=2)

                # EXIBE COM UNIDADE
                if item["tipo_unidade"] == "Kg":
                    texto_qtd = f"{item['quantidade']:.2f} kg".replace(".00", "")
                else:
                    texto_qtd = f"{int(item['quantidade'])} unid." if item['quantidade'] == int(item['quantidade']) else f"{item['quantidade']} unid."
                tk.Label(frame_item, text=f"Qtd: {texto_qtd}", font=("Arial", 11, "bold"),
                         fg="white", bg="#28a745" if item["quantidade"] > 0 else "#dc3545", padx=5, pady=3).pack(pady=2)

                fb = tk.Frame(frame_item, bg=cor)
                fb.pack(pady=5, anchor="center")
                tk.Entry(fb, textvariable=item["var_esq"], width=4, justify="center", font=("Arial", 10)).pack(side="left", padx=2)
                tk.Button(fb, text="➖", font=("Arial", 12, "bold"), bg="#dc3545", fg="white",
                          command=lambda i=item: subtrair_quantidade(i), width=3, height=1).pack(side="left", padx=2)
                tk.Button(fb, text="➕", font=("Arial", 12, "bold"), bg="#28a745", fg="white",
                          command=lambda i=item: adicionar_quantidade(i), width=3, height=1).pack(side="left", padx=2)
                tk.Entry(fb, textvariable=item["var_dir"], width=4, justify="center", font=("Arial", 10)).pack(side="left", padx=2)

                tk.Label(frame_item, text=f"Criado: {item['data_criacao'][:10]}", font=("Arial", 8), fg="#888", bg=cor).pack(pady=2)

                def click(i=item): selecionar_item(i)
                frame_item.bind("<Button-1>", lambda e, i=item: click(i))
                for child in frame_item.winfo_children():
                    child.bind("<Button-1>", lambda e, i=item: click(i))

                col += 1
                if col >= num_cols:
                    col = 0
                    row += 1
        row += 1

    for c in range(num_cols):
        scroll_frame.grid_columnconfigure(c, weight=1)
    scroll_frame.update_idletasks()
    conteudo_canvas.configure(scrollregion=conteudo_canvas.bbox("all"))
    updating = False


def abrir_excel():
    if os.path.exists(CAMINHO_DB):
        os.startfile(CAMINHO_DB)
    else:
        messagebox.showwarning("Abrir Excel", "Arquivo não encontrado!")


def alterar_imagem(item):
    if not open_password_form("Alterar Imagem"):
        return
    caminho = filedialog.askopenfilename(title="Escolha a nova foto", filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif")])
    if not caminho:
        return

    salvar_estado()
    if item.get("image_path") and os.path.exists(item["image_path"]):
        os.remove(item["image_path"])
    novo = salvar_imagem(caminho, item["nome"], item["id"])
    if novo:
        item["image_path"] = novo
        item["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        salvar_no_excel()
        registrar_historico(f"Imagem alterada para '{item['nome']}' (ID:{item['id']})")
        atualizar_tela()


# --------------------------  INTERFACE PRINCIPAL  --------------------------
container_principal = tk.Frame(root, bg="#111")
container_principal.pack(fill="both", expand=True)

barra_botoes = tk.Frame(container_principal, bg="#333", height=60)
barra_botoes.pack(fill="x")

def estilo_botao(master, texto, comando, cor_fundo, cor_hover, lado):
    btn = tk.Button(master, text=texto, font=("Arial", 11, "bold"),
                    bg=cor_fundo, fg="white", activebackground=cor_hover,
                    command=comando, relief="flat", bd=0, padx=10, pady=6)
    btn.pack(side=lado, padx=4, pady=4)
    def on_enter(e): btn.config(bg=cor_hover)
    def on_leave(e): btn.config(bg=cor_fundo)
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    return btn

# Botões da esquerda
estilo_botao(barra_botoes, "Adicionar", adicionar_item, "#28a745", "#218838", "left")
estilo_botao(barra_botoes, "Editar", editar_item, "#ffc107", "#e0a800", "left")
estilo_botao(barra_botoes, "Remover", remover_item, "#dc3545", "#c82333", "left")
estilo_botao(barra_botoes, "Restaurar", restaurar_item, "#17a2b8", "#138496", "left")
estilo_botao(barra_botoes, "Abrir Excel", abrir_excel, "#6c757d", "#5a6268", "left")

# Busca
frame_center = tk.Frame(barra_botoes, bg="#333")
frame_center.pack(side="left", fill="x", expand=True)
entry_search = tk.Entry(frame_center, textvariable=search_var, font=("Arial", 12), bg="#444", fg="#aaa", insertbackground="white")
entry_search.pack(pady=10, padx=20, fill="x")
entry_search.insert(0, placeholder_text)
entry_search.bind("<FocusIn>", on_search_focus_in)
entry_search.bind("<FocusOut>", on_search_focus_out)
entry_search.bind("<KeyRelease>", on_search_keyrelease)

# Botões da direita
estilo_botao(barra_botoes, "Voltar", undo, "#6c757d", "#5a6268", "right")
estilo_botao(barra_botoes, "Avançar", redo, "#6c757d", "#5a6268", "right")
estilo_botao(barra_botoes, "Exportar", exportar_estoque, "#17a2b8", "#138496", "right")
btn_toggle_historico = estilo_botao(barra_botoes, "Histórico", toggle_historico, "#6c757d", "#5a6268", "right")

# Área de conteúdo
frame_conteudo = tk.Frame(container_principal, bg="#111")
frame_conteudo.pack(side="left", fill="both", expand=True)

conteudo_canvas = tk.Canvas(frame_conteudo, bg="#111", highlightthickness=0)
conteudo_canvas.pack(side="left", fill="both", expand=True)
scrollbar = tk.Scrollbar(frame_conteudo, orient="vertical", command=conteudo_canvas.yview)
scrollbar.pack(side="right", fill="y")
conteudo_canvas.configure(yscrollcommand=scrollbar.set)

scroll_frame = tk.Frame(conteudo_canvas, bg="#111")
conteudo_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

def _on_mousewheel(event):
    conteudo_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
conteudo_canvas.bind("<MouseWheel>", _on_mousewheel)

# Painel de histórico
painel_historico = tk.Frame(container_principal, bg="#222", width=300)

# --------------------------  INICIALIZAÇÃO  --------------------------
root.bind("<Control-z>", undo)
root.bind("<Control-Shift-z>", redo)
root.bind("<Configure>", ajustar_modo_visualizacao)

carregar_do_excel()
for cat in categorias:
    categoria_aberta[cat] = True
ajustar_modo_visualizacao()
atualizar_tela()

root.mainloop()