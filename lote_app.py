from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from lote_controller import LoteController


class LoteApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.controller = LoteController()
        self.root.title("Gestão de Lotes Industriais")
        self._configurar_estilos()
        self._build_main()
        self._refresh_lotes()

    def _configurar_estilos(self):
        style = ttk.Style(self.root)
        style.configure("Primary.TButton", padding=(16, 10), font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", foreground=[("!disabled", "white")], background=[("!disabled", "#2563eb")])
        style.configure("Secondary.TButton", padding=(12, 10), font=("Segoe UI", 10, "bold"))
        style.map("Secondary.TButton", foreground=[("!disabled", "#374151")], background=[("!disabled", "#e5e7eb")])
        style.configure("App.TCombobox", padding=(6, 4))
        style.configure("App.TSpinbox", padding=(6, 4))
        style.configure("App.TEntry", padding=(6, 4))

    def _build_main(self):
        top = ttk.Frame(self.root)
        top.pack(fill="x")
        ttk.Button(top, text="Gerar Lote", style="Primary.TButton", command=self._abrir_gerar_lote).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="Abrir Lote", style="Secondary.TButton", command=self._abrir_lote).pack(side="left", padx=4)
        ttk.Button(top, text="Exibir Lote", style="Secondary.TButton", command=self._exibir_lote).pack(side="left", padx=4)
        ttk.Button(top, text="Relatório", style="Secondary.TButton", command=self._abrir_relatorio).pack(side="left", padx=4)
        self.tree = ttk.Treeview(self.root, columns=("ident", "criado", "nf", "qtd", "status"), show="headings")
        for c, t in [("ident", "IDENT"), ("criado", "Criação"), ("nf", "NF Ref."), ("qtd", "Qtd. itens"), ("status", "Status")]:
            self.tree.heading(c, text=t)
        self.tree.pack(fill="both", expand=True)
        self.status = ttk.Label(self.root, text="Pronto")
        self.status.pack(fill="x", side="bottom")

    def _refresh_lotes(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.lotes = self.controller.listar_lotes()
        for lote in self.lotes:
            self.tree.insert("", "end", iid=lote.id, values=(lote.ident, lote.criado_em[:19], lote.nf_referencia, len(lote.itens), lote.status))
        self.status.config(text=f"Total de lotes: {len(self.lotes)}")

    def _abrir_gerar_lote(self):
        caminho = filedialog.askopenfilename(filetypes=[("XML", "*.xml")])
        if not caminho:
            return
        try:
            itens = self.controller.importar_xml(caminho)
        except ValueError as exc:
            messagebox.showerror("Erro", str(exc))
            return
        lote = self.controller.criar_lote(1, 2026, itens[0].nf_numero, "122", itens)
        self.controller.reimprimir_lote(lote)
        self.controller.atualizar_status(lote.id, "impresso")
        self._refresh_lotes()
        messagebox.showinfo("Sucesso", f"Lote {lote.ident} criado e enviado para impressão.")

    def _abrir_lote(self):
        self._exibir_lote()

    def _exibir_lote(self):
        sel = self.tree.selection()
        if not sel:
            return
        lote = next((l for l in self.lotes if l.id == sel[0]), None)
        if not lote:
            return
        win = tk.Toplevel(self.root)
        win.title(f"Lote {lote.ident}")
        ttk.Label(win, text=f"Status: {lote.status}").pack(anchor="w")
        tv = ttk.Treeview(win, columns=("cod", "desc", "qty", "nf", "ger"), show="headings")
        for c in ("cod", "desc", "qty", "nf", "ger"):
            tv.heading(c, text=c)
        tv.pack(fill="both", expand=True)
        for i in lote.itens:
            tv.insert("", "end", values=(i.cod_item, i.descr_item, i.qty, i.nf_numero, i.cod_gerado))
        ttk.Button(win, text="Imprimir novamente", command=lambda: self.controller.reimprimir_lote(lote)).pack()

    def _abrir_relatorio(self):
        win = tk.Toplevel(self.root)
        win.title("Relatório")
        ttk.Label(win, text="Relatório de lotes").pack()
