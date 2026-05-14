from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

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

    def _selecionar_codigos_para_impressao(self, lote) -> list[str]:
        dialog = tk.Toplevel(self.root)
        dialog.title("Selecionar itens para impressão")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text=f"Lote {lote.ident} - selecione os itens:").pack(anchor="w", padx=10, pady=8)
        lb = tk.Listbox(dialog, selectmode=tk.MULTIPLE, width=80, height=12)
        lb.pack(fill="both", expand=True, padx=10)
        for idx, item in enumerate(lote.itens):
            lb.insert("end", f"{item.cod_gerado} | {item.cod_item} | {item.descr_item} | Qtd: {item.qty}")
            lb.selection_set(idx)

        selecionados: list[str] = []

        def confirmar():
            inds = lb.curselection()
            selecionados.extend([lote.itens[i].cod_gerado for i in inds])
            dialog.destroy()

        ttk.Button(dialog, text="Confirmar seleção", style="Primary.TButton", command=confirmar).pack(pady=10)
        self.root.wait_window(dialog)
        return selecionados

    def _abrir_gerar_lote(self):
        caminho = filedialog.askopenfilename(filetypes=[("XML", "*.xml")])
        if not caminho:
            return
        try:
            itens = self.controller.importar_xml(caminho)
            seq_lote = simpledialog.askinteger("Nº do lote", "Informe o número do lote:", parent=self.root, minvalue=1, maxvalue=99)
            if seq_lote is None:
                return
            mes_rec = simpledialog.askinteger("Mês", "Informe o mês do recebimento (1-12):", parent=self.root, minvalue=1, maxvalue=12)
            if mes_rec is None:
                return
            ano_rec = simpledialog.askinteger("Ano", "Informe o ano do recebimento (ex: 2026):", parent=self.root, minvalue=2020, maxvalue=2040)
            if ano_rec is None:
                return
            lote = self.controller.criar_lote(seq_lote, mes_rec, ano_rec, itens[0].nf_numero, itens)
            codigos = self._selecionar_codigos_para_impressao(lote)
            if not codigos:
                self.controller.atualizar_status(lote.id, "confirmado")
                self._refresh_lotes()
                messagebox.showwarning("Atenção", "Nenhum item selecionado para impressão.")
                return
            resultado = self.controller.reimprimir_lote(lote, codigos=codigos)
            self.controller.atualizar_status(lote.id, "impresso")
            self._refresh_lotes()
            if resultado["enviados"]:
                messagebox.showinfo("Sucesso", f"Lote {lote.ident} enviado para impressão.")
            else:
                messagebox.showwarning("Aviso", f"Impressão não enviada automaticamente. PNGs salvos em: {resultado['arquivos'][0].rsplit('/', 1)[0] if resultado['arquivos'] else 'n/a'}")
        except ValueError as exc:
            messagebox.showerror("Erro", str(exc))

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
