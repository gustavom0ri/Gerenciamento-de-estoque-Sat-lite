import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
from PIL import Image, ImageTk
import datetime
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import shutil
import copy

estoque = []
removidos = []
item_selecionado = None
historico = []
painel_aberto = False
undo_stack = []
redo_stack = []
ultima_data_registrada = None  # Nova variável para rastrear a data da última ação

PASTA_IMAGENS = "imagens_estoque"
CAMINHO_HISTORICO = "historico.txt"  # Arquivo fixo para o histórico
if not os.path.exists(PASTA_IMAGENS):
    os.makedirs(PASTA_IMAGENS)

CAMINHO_DB = "estoque.xlsx"
TAMANHO_CARD = 250

root = tk.Tk()
root.title("Sistema de Estoque Satelite")
root.state("zoomed")


# ---- BANCO DE DADOS ----
def salvar_no_excel():
    try:
        with pd.ExcelWriter(CAMINHO_DB, engine='openpyxl') as writer:
            if estoque:
                dados_estoque = []
                for item in estoque:
                    dados_estoque.append({
                        "Nome": item["nome"],
                        "Id": item.get("id", f"ID_{len(dados_estoque) + 1}"),
                        "Quantidade": item["quantidade"],
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
                        "Data_Criacao": item.get("data_criacao", ""),
                        "Data_Alteracao": item.get("data_alteracao", ""),
                        "Data_Remocao": item.get("data_remocao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
                    })
                df_removidos = pd.DataFrame(dados_removidos)
                df_removidos.to_excel(writer, sheet_name='Removidos', index=False)
        print(f"Excel salvo com sucesso: {CAMINHO_DB}")
    except Exception as e:
        print(f"Erro ao salvar Excel: {e}")


def carregar_do_excel():
    global estoque, removidos
    estoque.clear()
    removidos.clear()

    if not os.path.exists(CAMINHO_DB):
        print(f"Arquivo {CAMINHO_DB} não encontrado. Iniciando com estoque vazio.")
        return

    try:
        with pd.ExcelFile(CAMINHO_DB) as xls:
            if 'Estoque' in xls.sheet_names:
                df_estoque = pd.read_excel(xls, 'Estoque')
                print(f"Carregando {len(df_estoque)} itens da aba 'Estoque'...")
                for _, row in df_estoque.iterrows():
                    nome_arquivo = f"{row['Nome']}_{row.get('Id', 'default')}.jpg"
                    caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)

                    imagem = None
                    if os.path.exists(caminho_imagem):
                        try:
                            img = Image.open(caminho_imagem)
                            img.thumbnail((TAMANHO_CARD - 20, TAMANHO_CARD - 100))
                            imagem = ImageTk.PhotoImage(img)
                            print(f"Imagem carregada: {caminho_imagem}")
                        except Exception as e:
                            print(f"Erro ao carregar imagem {caminho_imagem}: {e}")

                    item = {
                        "imagem": imagem,
                        "nome": str(row["Nome"]),
                        "quantidade": int(row["Quantidade"]),
                        "var_esq": tk.IntVar(value=1),
                        "var_dir": tk.IntVar(value=1),
                        "id": str(row.get("Id", f"ID_{len(estoque) + 1}")),
                        "data_criacao": str(
                            row.get("Data_Criacao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))),
                        "data_alteracao": str(
                            row.get("Data_Alteracao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))
                    }
                    estoque.append(item)
                    print(f"Item carregado: {item['nome']} (ID: {item['id']}, Qtd: {item['quantidade']})")

            if 'Removidos' in xls.sheet_names:
                df_removidos = pd.read_excel(xls, 'Removidos')
                print(f"Carregando {len(df_removidos)} itens da aba 'Removidos'...")
                for _, row in df_removidos.iterrows():
                    item = {
                        "nome": str(row["Nome"]),
                        "quantidade": int(row["Quantidade"]),
                        "id": str(row.get("Id", "")),
                        "data_criacao": str(row.get("Data_Criacao", "")),
                        "data_alteracao": str(row.get("Data_Alteracao", "")),
                        "data_remocao": str(row.get("Data_Remocao", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))
                    }
                    removidos.append(item)
                    print(f"Item removido carregado: {item['nome']} (ID: {item['id']})")
        print(f"Carregamento concluído: {len(estoque)} itens no estoque, {len(removidos)} itens removidos.")
    except Exception as e:
        print(f"Erro ao carregar Excel: {e}")
        messagebox.showerror("Erro", f"Falha ao carregar o arquivo de estoque: {e}")


def salvar_imagem(caminho_original, nome_item, item_id):
    if not os.path.exists(caminho_original):
        print(f"Imagem não encontrada: {caminho_original}")
        return None

    nome_arquivo = f"{nome_item}_{item_id}.jpg"
    caminho_destino = os.path.join(PASTA_IMAGENS, nome_arquivo)

    try:
        shutil.copy2(caminho_original, caminho_destino)
        img = Image.open(caminho_destino)
        img.thumbnail((TAMANHO_CARD - 20, TAMANHO_CARD - 100))
        print(f"Imagem salva: {caminho_destino}")
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"Erro ao salvar imagem: {e}")
        return None


# ---- UNDO/REDO ----
def salvar_estado():
    global estoque, removidos
    estado = {
        'estoque': [],
        'removidos': copy.deepcopy(removidos)
    }

    for item in estoque:
        estado_estoque = {
            'nome': item['nome'],
            'quantidade': item['quantidade'],
            'id': item['id'],
            'data_criacao': item['data_criacao'],
            'data_alteracao': item['data_alteracao']
        }
        estado['estoque'].append(estado_estoque)

    undo_stack.append(estado)
    if len(undo_stack) > 20:
        undo_stack.pop(0)
    print(f"Estado salvo: {len(estado['estoque'])} itens no estoque, redo_stack size: {len(redo_stack)}")


def recarregar_imagens_estoque(estoque_dados):
    novo_estoque = []
    for item_data in estoque_dados:
        nome_arquivo = f"{item_data['nome']}_{item_data['id']}.jpg"
        caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)

        imagem = None
        if os.path.exists(caminho_imagem):
            try:
                img = Image.open(caminho_imagem)
                img.thumbnail((TAMANHO_CARD - 20, TAMANHO_CARD - 100))
                imagem = ImageTk.PhotoImage(img)
            except Exception:
                pass

        item = {
            "imagem": imagem,
            "nome": item_data['nome'],
            "quantidade": item_data['quantidade'],
            "var_esq": tk.IntVar(value=1),
            "var_dir": tk.IntVar(value=1),
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

    estado_atual = {
        'estoque': [],
        'removidos': copy.deepcopy(removidos)
    }
    for item in estoque:
        estado_estoque = {
            'nome': item['nome'],
            'quantidade': item['quantidade'],
            'id': item['id'],
            'data_criacao': item['data_criacao'],
            'data_alteracao': item['data_alteracao']
        }
        estado_atual['estoque'].append(estado_estoque)
    redo_stack.append(estado_atual)
    if len(redo_stack) > 20:
        redo_stack.pop(0)
    print(f"Undo executado, redo_stack size: {len(redo_stack)}")

    estado_anterior = undo_stack.pop()
    removidos = copy.deepcopy(estado_anterior['removidos'])
    estoque = recarregar_imagens_estoque(estado_anterior['estoque'])

    salvar_no_excel()
    registrar_historico("↩️ Desfeita a última ação")
    atualizar_tela()


def redo(event=None):
    global estoque, removidos
    if not redo_stack:
        messagebox.showinfo("Redo", "Nada para refazer!")
        return

    estado_atual = {
        'estoque': [],
        'removidos': copy.deepcopy(removidos)
    }
    for item in estoque:
        estado_estoque = {
            'nome': item['nome'],
            'quantidade': item['quantidade'],
            'id': item['id'],
            'data_criacao': item['data_criacao'],
            'data_alteracao': item['data_alteracao']
        }
        estado_atual['estoque'].append(estado_estoque)
    undo_stack.append(estado_atual)
    if len(undo_stack) > 20:
        undo_stack.pop(0)

    estado_redo = redo_stack.pop()
    removidos = copy.deepcopy(estado_redo['removidos'])
    estoque = recarregar_imagens_estoque(estado_redo['estoque'])
    print(f"Redo executado, redo_stack size: {len(redo_stack)}, undo_stack size: {len(undo_stack)}")

    salvar_no_excel()
    registrar_historico("↪️ Refeita a ação desfeita")
    atualizar_tela()


# ---- HISTÓRICO ----
def exportar_historico():
    if not historico:
        messagebox.showwarning("Exportar", "Histórico vazio!")
        return

    caminho = filedialog.asksaveasfilename(
        title="Salvar Histórico",
        defaultextension=".txt",
        filetypes=[("Arquivo de Texto", "*.txt")]
    )
    if not caminho:
        return

    try:
        with open(caminho, "w", encoding="utf-8") as arquivo:
            for acao in historico:
                arquivo.write(acao + "\n")
        messagebox.showinfo("Exportar", f"Histórico exportado com sucesso em:\n{caminho}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao exportar histórico:\n{e}")


def registrar_historico(acao):
    global ultima_data_registrada
    agora = datetime.datetime.now()
    data_atual = agora.strftime("%d/%m/%Y")
    hora = agora.strftime("%H:%M:%S")
    acao_completa = f"[{hora}] {acao}"

    # Verificar mudança de dia
    if ultima_data_registrada and ultima_data_registrada != data_atual:
        historico.append("--------")
        try:
            with open(CAMINHO_HISTORICO, "a", encoding="utf-8") as arquivo:
                arquivo.write("--------\n")
        except Exception as e:
            print(f"Erro ao gravar separador no arquivo {CAMINHO_HISTORICO}: {e}")

    historico.append(acao_completa)
    ultima_data_registrada = data_atual

    # Gravar no arquivo historico.txt
    try:
        with open(CAMINHO_HISTORICO, "a", encoding="utf-8") as arquivo:
            arquivo.write(acao_completa + "\n")
        print(f"Histórico gravado em {CAMINHO_HISTORICO}: {acao_completa}")
    except Exception as e:
        print(f"Erro ao gravar histórico em {CAMINHO_HISTORICO}: {e}")

    if painel_aberto:
        atualizar_historico()


def toggle_historico():
    global painel_aberto
    if painel_aberto:
        painel_historico.pack_forget()
        painel_aberto = False
    else:
        painel_historico.pack(side="right", fill="y", padx=(0, 0))
        painel_aberto = True
        atualizar_historico()
    atualizar_tela()


def atualizar_historico():
    for widget in painel_historico.winfo_children():
        widget.destroy()

    if not historico:
        return

    tk.Label(
        painel_historico,
        text="📜 Histórico",
        font=("Arial", 14, "bold"),
        bg="#222",
        fg="white"
    ).pack(pady=5)

    frame_principal = tk.Frame(painel_historico, bg="#222")
    frame_principal.pack(fill="both", expand=True)

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

    def _on_mousewheel(event):
        canvas_hist.yview_scroll(int(-1 * (event.delta / 120)), "units")

    painel_historico.bind_all("<MouseWheel>", _on_mousewheel)

    for acao in historico[::-1]:
        tk.Label(
            scroll_frame_hist,
            text=acao,
            anchor="w",
            justify="left",
            wraplength=250,
            font=("Arial", 11),
            bg="#333" if acao != "--------" else "#222",
            fg="white" if acao != "--------" else "#888"
        ).pack(fill="x", padx=2, pady=1)

    btn_export = tk.Button(
        painel_historico,
        text="💾 Exportar Log",
        command=exportar_historico,
        bg="#555",
        fg="white",
        font=("Arial", 12)
    )
    btn_export.pack(side="bottom", pady=5, padx=5, fill="x")


# ---- ITENS ----
def adicionar_item():
    global item_selecionado
    caminho_imagem = filedialog.askopenfilename(
        title="Escolha a foto do item",
        filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif")],
        parent=root
    )
    if not caminho_imagem:
        return

    nome = simpledialog.askstring("Nome do Item", "Digite o nome do item:", parent=root)
    if not nome:
        return

    quantidade = None

    def confirmar(event=None):
        nonlocal quantidade
        try:
            quantidade = int(entry.get())
            topo.destroy()
        except ValueError:
            messagebox.showerror("Erro", "Digite um número válido!")

    topo = tk.Toplevel(root)
    topo.title("Quantidade")
    tk.Label(topo, text=f"Digite a quantidade para '{nome}':").pack(padx=10, pady=10)
    entry = tk.Entry(topo)
    entry.insert(0, "1")
    entry.pack(padx=10, pady=5)
    entry.focus_set()
    tk.Button(topo, text="OK", command=confirmar).pack(pady=10)
    topo.bind("<Return>", confirmar)
    topo.transient(root)
    topo.grab_set()
    root.wait_window(topo)

    if quantidade is None:
        return

    salvar_estado()
    item_id = f"ID_{len(estoque) + 1}"
    data_agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    imagem_tk = salvar_imagem(caminho_imagem, nome, item_id)
    if not imagem_tk:
        messagebox.showerror("Erro", "Falha ao processar a imagem!")
        return

    item = {
        "imagem": imagem_tk,
        "nome": nome,
        "quantidade": quantidade,
        "var_esq": tk.IntVar(value=1),
        "var_dir": tk.IntVar(value=1),
        "id": item_id,
        "data_criacao": data_agora,
        "data_alteracao": data_agora
    }
    estoque.append(item)
    salvar_no_excel()
    registrar_historico(f"➕ Adicionado '{nome}' (ID: {item_id}) com {quantidade} unidades")
    atualizar_tela()


def remover_item():
    global item_selecionado
    if not item_selecionado:
        messagebox.showwarning("Remover", "Nenhum item selecionado!")
        return

    salvar_estado()
    item_removido = {
        "nome": item_selecionado['nome'],
        "quantidade": item_selecionado['quantidade'],
        "id": item_selecionado['id'],
        "data_criacao": item_selecionado['data_criacao'],
        "data_alteracao": item_selecionado['data_alteracao'],
        "data_remocao": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    removidos.append(item_removido)
    registrar_historico(
        f"➖ Removido '{item_selecionado['nome']}' (ID: {item_selecionado['id']}) - Movido para Removidos")
    estoque.remove(item_selecionado)
    item_selecionado = None
    salvar_no_excel()
    atualizar_tela()


def restaurar_item():
    global removidos
    if not removidos:
        messagebox.showinfo("Restaurar", "Nenhum item removido para restaurar!")
        return

    topo = tk.Toplevel(root)
    topo.title("Restaurar Item")
    topo.geometry("450x350")
    topo.transient(root)
    topo.grab_set()

    tk.Label(topo, text="Selecione o item para restaurar:", font=("Arial", 12, "bold")).pack(pady=10)

    listbox = tk.Listbox(topo, font=("Arial", 10), height=12)
    for idx, item in enumerate(removidos):
        texto = f"{item['id']} - {item['nome']} (Qtd: {item['quantidade']})"
        texto += f"\n   Removido: {item.get('data_remocao', 'N/A')}"
        listbox.insert(tk.END, texto)
    listbox.pack(fill="both", expand=True, padx=10, pady=10)

    def confirmar():
        selecao = listbox.curselection()
        if not selecao:
            messagebox.showwarning("Seleção", "Selecione um item para restaurar!")
            return

        idx = selecao[0]
        item_restaurar = removidos.pop(idx)

        nome_arquivo = f"{item_restaurar['nome']}_{item_restaurar['id']}.jpg"
        caminho_imagem = os.path.join(PASTA_IMAGENS, nome_arquivo)
        imagem = None
        if os.path.exists(caminho_imagem):
            try:
                img = Image.open(caminho_imagem)
                img.thumbnail((TAMANHO_CARD - 20, TAMANHO_CARD - 100))
                imagem = ImageTk.PhotoImage(img)
            except Exception:
                pass

        novo_item = {
            "imagem": imagem,
            "nome": item_restaurar["nome"],
            "quantidade": item_restaurar["quantidade"],
            "var_esq": tk.IntVar(value=1),
            "var_dir": tk.IntVar(value=1),
            "id": item_restaurar["id"],
            "data_criacao": item_restaurar["data_criacao"],
            "data_alteracao": item_restaurar["data_alteracao"]
        }

        salvar_estado()
        estoque.append(novo_item)
        salvar_no_excel()
        registrar_historico(f"🔄 Restaurado '{novo_item['nome']}' (ID: {novo_item['id']}) dos Removidos")
        atualizar_tela()
        topo.destroy()

    def cancelar():
        topo.destroy()

    frame_botoes = tk.Frame(topo)
    frame_botoes.pack(pady=10)

    tk.Button(frame_botoes, text="Restaurar", command=confirmar, bg="#28a745", fg="white",
              font=("Arial", 10, "bold")).pack(side="left", padx=5)
    tk.Button(frame_botoes, text="Cancelar", command=cancelar, bg="#6c757d", fg="white",
              font=("Arial", 10, "bold")).pack(side="left", padx=5)


def editar_item():
    global item_selecionado
    if not item_selecionado:
        messagebox.showwarning("Editar", "Nenhum item selecionado!")
        return

    salvar_estado()
    nome_antigo = item_selecionado["nome"]
    qtd_antiga = item_selecionado["quantidade"]
    id_item = item_selecionado["id"]

    novo_nome = simpledialog.askstring(
        "Editar Item",
        "Digite o novo nome do item:",
        initialvalue=item_selecionado["nome"],
        parent=root
    )
    if novo_nome is None:
        return
    if not novo_nome:
        messagebox.showerror("Erro", "O nome não pode estar vazio!")
        return
    item_selecionado["nome"] = novo_nome

    nova_qtd = None

    def confirmar(event=None):
        nonlocal nova_qtd
        try:
            nova_qtd = int(entry.get())
            topo.destroy()
        except ValueError:
            messagebox.showerror("Erro", "Digite um número válido!")

    topo = tk.Toplevel(root)
    topo.title("Editar Quantidade")
    tk.Label(topo, text=f"Digite a nova quantidade para '{novo_nome}':").pack(padx=10, pady=10)
    entry = tk.Entry(topo)
    entry.insert(0, str(item_selecionado["quantidade"]))
    entry.pack(padx=10, pady=5)
    entry.focus_set()
    tk.Button(topo, text="OK", command=confirmar).pack(pady=10)
    topo.bind("<Return>", confirmar)
    topo.transient(root)
    topo.grab_set()
    root.wait_window(topo)

    if nova_qtd is not None:
        item_selecionado["quantidade"] = nova_qtd
        item_selecionado["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        if novo_nome != nome_antigo:
            nome_arquivo_antigo = f"{nome_antigo}_{id_item}.jpg"
            nome_arquivo_novo = f"{novo_nome}_{id_item}.jpg"
            caminho_antigo = os.path.join(PASTA_IMAGENS, nome_arquivo_antigo)
            caminho_novo = os.path.join(PASTA_IMAGENS, nome_arquivo_novo)

            if os.path.exists(caminho_antigo):
                try:
                    shutil.move(caminho_antigo, caminho_novo)
                    img = Image.open(caminho_novo)
                    img.thumbnail((TAMANHO_CARD - 20, TAMANHO_CARD - 100))
                    item_selecionado["imagem"] = ImageTk.PhotoImage(img)
                except Exception:
                    pass

        salvar_no_excel()
        registrar_historico(f"✏️ Editado '{nome_antigo}' (ID: {id_item}) ({qtd_antiga}) → '{novo_nome}' ({nova_qtd})")
        atualizar_tela()


def adicionar_quantidade(item):
    try:
        valor = item["var_dir"].get()
        if valor <= 0:
            messagebox.showwarning("Aviso", "Digite um número maior que 0!")
            return

        salvar_estado()
        item["quantidade"] += valor
        item["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        item["var_dir"].set(1)
        salvar_no_excel()
        registrar_historico(
            f"⬆️ Adicionado +{valor} em '{item['nome']}' (ID: {item['id']}) → total {item['quantidade']}")
        atualizar_tela()
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido!")


def subtrair_quantidade(item):
    try:
        valor = item["var_esq"].get()
        if valor <= 0:
            messagebox.showwarning("Aviso", "Digite um número maior que 0!")
            return

        salvar_estado()
        nova_qtd = max(0, item["quantidade"] - valor)
        item["quantidade"] = nova_qtd
        item["data_alteracao"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        item["var_esq"].set(1)
        salvar_no_excel()
        registrar_historico(f"⬇️ Removido -{valor} de '{item['nome']}' (ID: {item['id']}) → total {nova_qtd}")
        atualizar_tela()
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido!")


def selecionar_item(item):
    global item_selecionado
    if item_selecionado == item:
        item_selecionado = None
    else:
        item_selecionado = item
    atualizar_tela()


def exportar_estoque():
    if not estoque:
        messagebox.showwarning("Exportar", "Nenhum item no estoque!")
        return

    df = pd.DataFrame([{
        "ID": item["id"],
        "Nome": item["nome"],
        "Quantidade": item["quantidade"],
        "Data_Criacao": item["data_criacao"],
        "Data_Alteracao": item["data_alteracao"]
    } for item in estoque])

    menu = tk.Toplevel(root)
    menu.title("Escolher formato")
    menu.geometry("300x200")
    menu.resizable(False, False)

    def salvar_excel():
        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if caminho:
            df.to_excel(caminho, index=False)
            messagebox.showinfo("Exportar", "Estoque exportado com sucesso!")
        menu.destroy()

    def salvar_csv():
        caminho = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if caminho:
            df.to_csv(caminho, index=False)
            messagebox.showinfo("Exportar", "Estoque exportado com sucesso!")
        menu.destroy()

    def salvar_txt():
        caminho = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("TXT", "*.txt")])
        if caminho:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write("ESTOQUE SATELITE\n")
                f.write("=" * 50 + "\n\n")
                for item in estoque:
                    f.write(f"ID: {item['id']}\n")
                    f.write(f"Nome: {item['nome']}\n")
                    f.write(f"Quantidade: {item['quantidade']}\n")
                    f.write(f"Data Criação: {item['data_criacao']}\n")
                    f.write(f"Data Alteração: {item['data_alteracao']}\n")
                    f.write("-" * 30 + "\n")
            messagebox.showinfo("Exportar", "Estoque exportado com sucesso!")
        menu.destroy()

    def salvar_pdf():
        caminho = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if caminho:
            c = canvas.Canvas(caminho, pagesize=letter)
            largura, altura = letter
            y = altura - 50
            c.setFont("Helvetica-Bold", 14)
            c.drawString(50, y, "Estoque Satelite")
            y -= 30
            c.setFont("Helvetica", 10)
            for i, item in enumerate(estoque, start=1):
                texto = f"{i}. {item['id']} - {item['nome']}"
                c.drawString(50, y, texto)
                y -= 15
                c.drawString(50, y, f"   Qtd: {item['quantidade']} | Criado: {item['data_criacao'][:10]}")
                y -= 15
                c.drawString(50, y, f"   Última alteração: {item['data_alteracao'][:10]}")
                y -= 20
                if y < 80:
                    c.showPage()
                    y = altura - 50
            c.save()
            messagebox.showinfo("Exportar", "Estoque exportado com sucesso!")
        menu.destroy()

    tk.Button(menu, text="📄 Excel (.xlsx)", width=25, command=salvar_excel).pack(pady=10)
    tk.Button(menu, text="📄 CSV (.csv)", width=25, command=salvar_csv).pack(pady=10)
    tk.Button(menu, text="📄 TXT (.txt)", width=25, command=salvar_txt).pack(pady=10)
    tk.Button(menu, text="📄 PDF (.pdf)", width=25, command=salvar_pdf).pack(pady=10)


def atualizar_tela():
    global item_selecionado
    print("Atualizando tela...")
    for widget in scroll_frame.winfo_children():
        widget.destroy()

    if not estoque:
        print("Estoque vazio, exibindo mensagem de estoque vazio.")
        label_vazio = tk.Label(
            scroll_frame,
            text="📦 Estoque vazio\nClique em 'Adicionar' para começar",
            font=("Arial", 16),
            fg="#888",
            bg="#111"
        )
        label_vazio.pack(expand=True, fill="both")
        conteudo_canvas.configure(scrollregion=conteudo_canvas.bbox("all"))
        return

    num_itens_por_linha = 4 if painel_aberto else 5
    print(f"Painel histórico aberto: {painel_aberto}, Itens por linha: {num_itens_por_linha}")

    largura_janela = frame_conteudo.winfo_width() or root.winfo_width()
    if painel_aberto:
        largura_janela -= 300  # Subtrai a largura do painel de histórico
    largura_item = largura_janela // num_itens_por_linha

    row = 0
    col = 0
    for item in estoque:
        print(f"Renderizando item: {item['nome']} (ID: {item['id']})")
        frame_principal = tk.Frame(scroll_frame, width=largura_item, height=TAMANHO_CARD + 50, bg="#111")
        frame_principal.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        frame_principal.grid_propagate(False)

        cor_card = "#2c2c2c"
        if item_selecionado == item:
            cor_card = "#4444aa"

        frame_item = tk.Frame(frame_principal, bg=cor_card, bd=3, relief="ridge")
        frame_item.pack(expand=True, fill="both", padx=5, pady=5)

        if item.get("imagem"):
            lbl_img = tk.Label(frame_item, image=item["imagem"], bg=cor_card)
            lbl_img.image = item["imagem"]
            lbl_img.pack(pady=5)

        lbl_nome = tk.Label(
            frame_item,
            text=f"{item['id']}\n{item['nome']}",
            font=("Arial", 10, "bold"),
            fg="#00d4ff",
            bg=cor_card,
            justify="center"
        )
        lbl_nome.pack(pady=2)

        lbl_qtd = tk.Label(
            frame_item,
            text=f"Qtd: {item['quantidade']}",
            font=("Arial", 11, "bold"),
            fg="white",
            bg="#28a745" if item["quantidade"] > 0 else "#dc3545",
            padx=5,
            pady=3
        )
        lbl_qtd.pack(pady=2)

        frame_botoes = tk.Frame(frame_item, bg=cor_card)
        frame_botoes.pack(fill="x", pady=10)

        btn_sub = tk.Button(
            frame_botoes,
            text="➖",
            font=("Arial", 16, "bold"),
            bg="#dc3545",
            fg="white",
            activebackground="#a71d2a",
            activeforeground="white",
            command=lambda i=item: subtrair_quantidade(i),
            relief="flat",
            width=5,
            height=1
        )
        btn_sub.pack(side="left", padx=10)

        entry_sub = tk.Entry(frame_botoes, textvariable=item["var_esq"], width=6, justify="center", font=("Arial", 12))
        entry_sub.pack(side="left", padx=5)

        btn_add = tk.Button(
            frame_botoes,
            text="➕",
            font=("Arial", 16, "bold"),
            bg="#28a745",
            fg="white",
            activebackground="#1e7e34",
            activeforeground="white",
            command=lambda i=item: adicionar_quantidade(i),
            relief="flat",
            width=5,
            height=1
        )
        btn_add.pack(side="left", padx=10)

        entry_add = tk.Entry(frame_botoes, textvariable=item["var_dir"], width=6, justify="center", font=("Arial", 12))
        entry_add.pack(side="left", padx=5)

        lbl_datas = tk.Label(
            frame_item,
            text=f"Criado: {item['data_criacao'][:10]}",
            font=("Arial", 8),
            fg="#888",
            bg=cor_card,
            justify="center"
        )
        lbl_datas.pack(pady=5)

        def on_click(i=item):
            selecionar_item(i)

        frame_item.bind("<Button-1>", lambda e, i=item: on_click(i))
        for child in frame_item.winfo_children():
            child.bind("<Button-1>", lambda e, i=item: on_click(i))

        col += 1
        if col >= num_itens_por_linha:
            col = 0
            row += 1

    for c in range(num_itens_por_linha):
        scroll_frame.grid_columnconfigure(c, weight=1)
    for r in range(row + 1):
        scroll_frame.grid_rowconfigure(r, weight=1)

    scroll_frame.update_idletasks()
    conteudo_canvas.configure(scrollregion=conteudo_canvas.bbox("all"))
    print("Tela atualizada com sucesso.")


def abrir_excel():
    if os.path.exists(CAMINHO_DB):
        os.startfile(CAMINHO_DB)
    else:
        messagebox.showwarning("Abrir Excel", "Arquivo de estoque não encontrado!")


# ---- INTERFACE ----
container_principal = tk.Frame(root, bg="#111")
container_principal.pack(fill="both", expand=True)

barra_botoes = tk.Frame(container_principal, bg="#333", height=60)
barra_botoes.pack(fill="x")


def estilo_botao(master, texto, comando, cor_fundo, cor_hover, lado):
    btn = tk.Button(
        master,
        text=texto,
        font=("Arial", 16, "bold"),
        bg=cor_fundo,
        fg="white",
        activebackground=cor_hover,
        activeforeground="white",
        command=comando,
        relief="flat",
        bd=0,
        padx=15,
        pady=8,
        highlightthickness=0
    )
    btn.pack(side=lado, padx=10, pady=10)

    def on_enter(e): btn.config(bg=cor_hover)

    def on_leave(e): btn.config(bg=cor_fundo)

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    return btn


btn_adicionar = estilo_botao(barra_botoes, "➕ Adicionar", adicionar_item, cor_fundo="#28a745", cor_hover="#218838",
                             lado="left")
btn_remover = estilo_botao(barra_botoes, "➖ Remover", remover_item, cor_fundo="#dc3545", cor_hover="#c82333",
                           lado="left")
btn_editar = estilo_botao(barra_botoes, "✏️ Editar", editar_item, cor_fundo="#ffc107", cor_hover="#e0a800", lado="left")
btn_restaurar = estilo_botao(barra_botoes, "🔄 Restaurar", restaurar_item, cor_fundo="#17a2b8", cor_hover="#138496",
                             lado="left")
btn_abrir_excel = estilo_botao(barra_botoes, "📊 Abrir Excel", abrir_excel, cor_fundo="#17a2b8", cor_hover="#138496",
                               lado="left")
btn_undo = estilo_botao(barra_botoes, "↩️ Voltar", undo, cor_fundo="#6c757d", cor_hover="#5a6268", lado="right")
btn_redo = estilo_botao(barra_botoes, "↪️ Avançar", redo, cor_fundo="#6c757d", cor_hover="#5a6268", lado="right")
btn_exportar = estilo_botao(barra_botoes, "📤 Exportar", exportar_estoque, cor_fundo="#17a2b8", cor_hover="#138496",
                            lado="right")
btn_historico = estilo_botao(barra_botoes, "📜 Histórico", toggle_historico, cor_fundo="#6c757d", cor_hover="#5a6268",
                             lado="right")

frame_conteudo = tk.Frame(container_principal, bg="#111")
frame_conteudo.pack(side="left", fill="both", expand=True)

conteudo_canvas = tk.Canvas(frame_conteudo, bg="#111", highlightthickness=0)
conteudo_canvas.pack(side="left", fill="both", expand=True, padx=(0, 10))

scrollbar = tk.Scrollbar(frame_conteudo, orient="vertical", command=conteudo_canvas.yview)
scrollbar.pack(side="right", fill="y")

conteudo_canvas.configure(yscrollcommand=scrollbar.set)

scroll_frame = tk.Frame(conteudo_canvas, bg="#111")
conteudo_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")


def _on_mousewheel(event):
    conteudo_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


frame_conteudo.bind_all("<MouseWheel>", _on_mousewheel)

scroll_frame.bind("<Configure>", lambda e: conteudo_canvas.configure(scrollregion=conteudo_canvas.bbox("all")))

painel_historico = tk.Frame(container_principal, bg="#222", width=300)

root.bind("<Control-z>", undo)
root.bind("<Control-Shift-z>", redo)
root.focus_set()

# ---- INICIALIZAÇÃO ----
print("Iniciando programa...")
carregar_do_excel()
atualizar_tela()
print("Programa iniciado.")
root.mainloop()