import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
from PIL import Image, ImageTk
import datetime
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

estoque = []
item_selecionado = None
historico = []  # Lista para armazenar as ações
painel_aberto = False  # Controle da barra lateral

# Caminho fixo para o "banco de dados" Excel
CAMINHO_DB = "estoque.xlsx"

root = tk.Tk()
root.title("Sistema de Estoque Satelite")
root.state("zoomed")


# ---- BANCO DE DADOS ----
def salvar_no_excel():
    """Salva o estoque no Excel como banco de dados"""
    if not estoque:
        return
    df = pd.DataFrame([{
        "Nome": item["nome"],
        "Quantidade": item["quantidade"]
    } for item in estoque])
    df.to_excel(CAMINHO_DB, index=False)


def carregar_do_excel():
    """Carrega os itens do Excel para o programa"""
    if not os.path.exists(CAMINHO_DB):
        return
    df = pd.read_excel(CAMINHO_DB)
    for _, row in df.iterrows():
        item = {
            "imagem": None,  # imagens não são salvas no Excel
            "nome": row["Nome"],
            "quantidade": int(row["Quantidade"]),
            "var_esq": tk.IntVar(value=1),
            "var_dir": tk.IntVar(value=1)
        }
        estoque.append(item)


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
    """Adiciona uma ação ao histórico com horário e atualiza a barra lateral, se aberta."""
    agora = datetime.datetime.now().strftime("%H:%M:%S")
    historico.append(f"[{agora}] {acao}")
    if painel_aberto:
        atualizar_historico()


def toggle_historico():
    global painel_aberto
    if painel_aberto:
        painel_historico.pack_forget()
        painel_aberto = False
    else:
        painel_historico.pack(side="right", fill="y")
        painel_aberto = True
        atualizar_historico()


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
    scroll_frame = tk.Frame(canvas_hist, bg="#222")

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas_hist.configure(scrollregion=canvas_hist.bbox("all"))
    )

    canvas_hist.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas_hist.configure(yscrollcommand=scrollbar.set)

    canvas_hist.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas_hist.yview_scroll(int(-1*(event.delta/120)), "units")

    painel_historico.bind_all("<MouseWheel>", _on_mousewheel)

    for acao in historico[::-1]:
        tk.Label(
            scroll_frame,
            text=acao,
            anchor="w",
            justify="left",
            wraplength=250,
            font=("Arial", 11),
            bg="#333",
            fg="white"
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

    imagem = Image.open(caminho_imagem)
    largura_tela = conteudo.winfo_width() or root.winfo_width()
    largura_item = largura_tela // 3 - 40
    imagem.thumbnail((largura_item, 250))
    imagem_tk = ImageTk.PhotoImage(imagem)

    item = {
        "imagem": imagem_tk,
        "nome": nome,
        "quantidade": quantidade,
        "var_esq": tk.IntVar(value=1),
        "var_dir": tk.IntVar(value=1)
    }
    estoque.append(item)
    salvar_no_excel()
    registrar_historico(f"➕ Adicionado '{nome}' com {quantidade} unidades")
    atualizar_tela()


def remover_item():
    global item_selecionado
    if item_selecionado:
        registrar_historico(f"➖ Removido '{item_selecionado['nome']}'")
        estoque.remove(item_selecionado)
        salvar_no_excel()
        item_selecionado = None
        atualizar_tela()
    else:
        messagebox.showwarning("Remover", "Nenhum item selecionado!")


def editar_item():
    global item_selecionado
    if not item_selecionado:
        messagebox.showwarning("Editar", "Nenhum item selecionado!")
        return

    nome_antigo = item_selecionado["nome"]
    qtd_antiga = item_selecionado["quantidade"]

    novo_nome = simpledialog.askstring(
        "Editar Item",
        "Digite o novo nome do item:",
        initialvalue=item_selecionado["nome"],
        parent=root
    )
    if not novo_nome:
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
        salvar_no_excel()
        registrar_historico(f"✏️ Editado '{nome_antigo}' ({qtd_antiga}) → '{novo_nome}' ({nova_qtd})")
        atualizar_tela()


def adicionar_quantidade(item):
    try:
        valor = int(item["var_dir"].get())
        item["quantidade"] += valor
        salvar_no_excel()
        registrar_historico(f"⬆️ Adicionado +{valor} em '{item['nome']}' → total {item['quantidade']}")
        atualizar_tela()
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido!")


def subtrair_quantidade(item):
    try:
        valor = int(item["var_esq"].get())
        nova_qtd = max(0, item["quantidade"] - valor)
        item["quantidade"] = nova_qtd
        salvar_no_excel()
        registrar_historico(f"⬇️ Removido -{valor} de '{item['nome']}' → total {nova_qtd}")
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
        "Nome": item["nome"],
        "Quantidade": item["quantidade"]
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
                for item in estoque:
                    f.write(f"{item['nome']} - Quantidade: {item['quantidade']}\n")
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
            c.setFont("Helvetica", 12)
            for i, item in enumerate(estoque, start=1):
                c.drawString(50, y, f"{i}. {item['nome']} - Quantidade: {item['quantidade']}")
                y -= 20
                if y < 50:
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
    for widget in conteudo.winfo_children():
        widget.destroy()

    largura_tela = conteudo.winfo_width() or root.winfo_width()
    largura_item = largura_tela // 3 - 40

    for idx, item in enumerate(estoque):
        linha = idx // 3
        coluna = idx % 3

        frame_principal = tk.Frame(conteudo, width=largura_item, bg="#111")
        frame_principal.grid(row=linha, column=coluna, padx=15, pady=15, sticky="nsew")
        frame_principal.grid_propagate(False)

        cor_card = "#2c2c2c"
        if item_selecionado == item:
            cor_card = "#4444aa"

        frame_item = tk.Frame(frame_principal, bg=cor_card, bd=3, relief="ridge")
        frame_item.pack(expand=True, fill="both", padx=5, pady=5)

        if item["imagem"]:
            lbl_img = tk.Label(frame_item, image=item["imagem"], bg=cor_card)
            lbl_img.image = item["imagem"]
            lbl_img.pack(pady=10)

        lbl_nome = tk.Label(
            frame_item,
            text=item["nome"],
            font=("Arial", 14, "bold"),
            fg="#00d4ff",
            bg=cor_card
        )
        lbl_nome.pack(pady=5)

        lbl_qtd = tk.Label(
            frame_item,
            text=f"Quantidade: {item['quantidade']}",
            font=("Arial", 12, "bold"),
            fg="white",
            bg="#28a745" if item["quantidade"] > 0 else "#dc3545",
            padx=10, pady=5
        )
        lbl_qtd.pack(pady=5, fill="x")

        frame_botoes = tk.Frame(frame_item, bg=cor_card)
        frame_botoes.pack(pady=10)

        def estilo_btn(master, texto, comando, cor, cor_hover):
            btn = tk.Button(
                master, text=texto, font=("Arial", 12, "bold"),
                bg=cor, fg="white", activebackground=cor_hover,
                activeforeground="white", command=comando,
                relief="flat", bd=0, padx=10, pady=5, width=3
            )
            btn.pack(side="left", padx=5)

            def on_enter(e): btn.config(bg=cor_hover)
            def on_leave(e): btn.config(bg=cor)
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            return btn

        entry_sub = tk.Entry(frame_botoes, textvariable=item["var_esq"], width=4, justify="center")
        entry_sub.pack(side="left", padx=2)
        estilo_btn(frame_botoes, "➖", lambda i=item: subtrair_quantidade(i), "#dc3545", "#a71d2a")

        estilo_btn(frame_botoes, "➕", lambda i=item: adicionar_quantidade(i), "#28a745", "#1e7e34")
        entry_add = tk.Entry(frame_botoes, textvariable=item["var_dir"], width=4, justify="center")
        entry_add.pack(side="left", padx=2)

        frame_item.bind("<Button-1>", lambda e, i=item: selecionar_item(i))
        for child in frame_item.winfo_children():
            child.bind("<Button-1>", lambda e, i=item: selecionar_item(i))

    for c in range(3):
        conteudo.grid_columnconfigure(c, weight=1)


# ---- INTERFACE ----
barra_botoes = tk.Frame(root, bg="#333", height=60)
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


btn_adicionar = estilo_botao(
    barra_botoes, "➕ Adicionar", adicionar_item,
    cor_fundo="#28a745", cor_hover="#218838", lado="left"
)

btn_remover = estilo_botao(
    barra_botoes, "➖ Remover", remover_item,
    cor_fundo="#dc3545", cor_hover="#c82333", lado="left"
)

btn_editar = estilo_botao(
    barra_botoes, "✏️ Editar", editar_item,
    cor_fundo="#ffc107", cor_hover="#e0a800", lado="left"
)

btn_exportar = estilo_botao(
    barra_botoes, "📤 Exportar", exportar_estoque,
    cor_fundo="#17a2b8", cor_hover="#138496", lado="right"
)

btn_historico = estilo_botao(
    barra_botoes, "📜 Histórico", toggle_historico,
    cor_fundo="#6c757d", cor_hover="#5a6268", lado="right"
)

conteudo = tk.Frame(root, bg="#111")
conteudo.pack(expand=True, fill="both")

painel_historico = tk.Frame(root, bg="#222", width=300)


# ---- INICIALIZAÇÃO ----
carregar_do_excel()
atualizar_tela()
root.mainloop()
