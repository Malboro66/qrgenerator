import io
import os
import queue
import tempfile
import zipfile
from dataclasses import dataclass

import pandas as pd
import qrcode
from PIL import Image, ImageTk
from qrcode.image.svg import SvgImage
from reportlab.graphics import renderPM
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


@dataclass
class ItemCodigo:
    """Representa um item que será convertido em código visual."""

    valor: str


class QRCodeGenerator:
    """Aplicativo desktop para geração de QR Codes e códigos de barras."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("QR / Código de Barras Generator")
        self.root.geometry("980x680")

        self.fila = queue.Queue()
        self.df = None
        self.arquivo_fonte = ""
        self.preview_image_ref = None

        self.qr_size = tk.IntVar(value=250)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.modo = tk.StringVar(value="texto")
        self.formato_saida = tk.StringVar(value="pdf")
        self.tipo_codigo = tk.StringVar(value="qrcode")

        self.prefixo_numerico = tk.StringVar(value="")
        self.sufixo_numerico = tk.StringVar(value="")

        self._criar_interface()
        self.atualizar_preview()

    def _criar_interface(self):
        topo = ttk.Frame(self.root, padding=10)
        topo.pack(fill="x")

        ttk.Button(topo, text="Selecionar Arquivo", command=self.selecionar_arquivo).grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )

        self.column_combo = ttk.Combobox(topo, state="disabled", width=35)
        self.column_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.generate_button = ttk.Button(
            topo,
            text="Gerar",
            state="disabled",
            command=self.gerar_a_partir_da_tabela,
        )
        self.generate_button.grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(topo, text="Tipo:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Radiobutton(
            topo,
            text="QR Code",
            variable=self.tipo_codigo,
            value="qrcode",
            command=self.atualizar_preview,
        ).grid(row=1, column=1, sticky="w")
        ttk.Radiobutton(
            topo,
            text="Código de Barras (Code128)",
            variable=self.tipo_codigo,
            value="barcode",
            command=self.atualizar_preview,
        ).grid(row=1, column=2, sticky="w")

        ttk.Label(topo, text="Modo de dados:").grid(row=2, column=0, sticky="w", padx=5)
        ttk.Radiobutton(
            topo,
            text="Texto",
            variable=self.modo,
            value="texto",
            command=self.atualizar_controles_formato,
        ).grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(
            topo,
            text="Numérico",
            variable=self.modo,
            value="numerico",
            command=self.atualizar_controles_formato,
        ).grid(row=2, column=2, sticky="w")

        self.texto_controls = ttk.Frame(topo)
        self.texto_controls.grid(row=3, column=0, columnspan=3, sticky="w", padx=5, pady=3)
        ttk.Label(self.texto_controls, text="Dados conforme coluna selecionada.").pack(anchor="w")

        self.numerico_controls = ttk.Frame(topo)
        ttk.Label(self.numerico_controls, text="Prefixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.prefixo_numerico, width=10).pack(
            side="left", padx=(0, 10)
        )
        ttk.Label(self.numerico_controls, text="Sufixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.sufixo_numerico, width=10).pack(side="left")

        self.atualizar_controles_formato()

        preview_frame = ttk.LabelFrame(self.root, text="Preview", padding=10)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.pack(expand=True)

    def atualizar_controles_formato(self):
        if self.modo.get() == "texto":
            self.numerico_controls.grid_forget()
            self.texto_controls.grid(row=3, column=0, columnspan=3, sticky="w", padx=5, pady=3)
        else:
            self.texto_controls.grid_forget()
            self.numerico_controls.grid(row=3, column=0, columnspan=3, sticky="w", padx=5, pady=3)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione CSV ou Excel",
            filetypes=[("Arquivos de dados", "*.csv *.xlsx")],
        )
        if not caminho:
            return

        try:
            if caminho.lower().endswith(".csv"):
                self.df = pd.read_csv(caminho)
            else:
                self.df = pd.read_excel(caminho)
        except Exception as exc:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {exc}")
            self.column_combo.configure(state="disabled", values=[])
            self.generate_button.configure(state="disabled")
            return

        self.arquivo_fonte = caminho
        colunas = list(self.df.columns)
        self.column_combo.configure(values=colunas, state="readonly")
        if colunas:
            self.column_combo.set(colunas[0])
            self.generate_button.configure(state="normal")
        else:
            self.generate_button.configure(state="disabled")

    def _normalizar_dado(self, valor: str) -> str:
        valor = str(valor)
        if self.modo.get() == "numerico":
            return f"{self.prefixo_numerico.get()}{valor}{self.sufixo_numerico.get()}"
        return valor

    def _gerar_imagem_obj(self, dado: str) -> Image.Image:
        if self.tipo_codigo.get() == "barcode":
            barcode = code128.Code128(dado, barHeight=20 * mm, barWidth=0.45)
            desenho = Drawing(self.qr_size.get(), int(self.qr_size.get() * 0.45))
            desenho.add(barcode)
            pil_image = renderPM.drawToPIL(desenho, dpi=200)
            return pil_image.convert("RGB")

        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(dado)
        qr.make(fit=True)
        img = qr.make_image(
            fill_color=self.qr_foreground_color.get(),
            back_color=self.qr_background_color.get(),
        )
        if hasattr(img, "get_image"):
            img = img.get_image()
        if not isinstance(img, Image.Image):
            img = Image.open(io.BytesIO(img.tobytes()))
        return img.convert("RGB")

    def atualizar_preview(self):
        amostra = self._normalizar_dado("123456789")
        img = self._gerar_imagem_obj(amostra)
        img.thumbnail((260, 260))
        self.preview_image_ref = ImageTk.PhotoImage(img)
        self.preview_label.configure(image=self.preview_image_ref)

    def gerar_imagens(self, codigos, formato, destino):
        os.makedirs(destino, exist_ok=True)
        total = len(codigos)

        for i, codigo in enumerate(codigos, start=1):
            dado = self._normalizar_dado(codigo)
            if formato == "svg":
                if self.tipo_codigo.get() == "barcode":
                    raise ValueError("Exportação SVG para código de barras não suportada nesta versão.")
                qr = qrcode.make(dado, image_factory=SvgImage)
                caminho_saida = os.path.join(destino, f"{codigo}.svg")
                with open(caminho_saida, "wb") as f:
                    qr.save(f)
            else:
                imagem = self._gerar_imagem_obj(dado)
                imagem.save(os.path.join(destino, f"{codigo}.png"), format="PNG")

            self.fila.put({"tipo": "progresso", "atual": i, "total": total, "codigo": codigo})

        self.fila.put({"tipo": "sucesso", "caminho": destino})

    def gerar_zip(self, codigos, caminho_zip):
        with tempfile.TemporaryDirectory() as tmpdir:
            self.gerar_imagens(codigos, "png", tmpdir)
            with zipfile.ZipFile(caminho_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for codigo in codigos:
                    nome = f"{codigo}.png"
                    zf.write(os.path.join(tmpdir, nome), arcname=nome)

    def gerar_pdf(self, codigos, caminho_pdf):
        pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
        largura_pagina, altura_pagina = A4

        x = 20 * mm
        y = altura_pagina - 40 * mm
        tamanho = 35 * mm
        margem = 10 * mm

        for codigo in codigos:
            imagem = self._gerar_imagem_obj(self._normalizar_dado(codigo))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                temp_path = tmp.name
                imagem.save(temp_path, format="PNG")

            pdf.drawImage(temp_path, x, y, width=tamanho, height=tamanho, preserveAspectRatio=True)
            os.unlink(temp_path)

            x += tamanho + margem
            if x + tamanho > largura_pagina - 20 * mm:
                x = 20 * mm
                y -= tamanho + margem

            if y < 20 * mm:
                pdf.showPage()
                x = 20 * mm
                y = altura_pagina - 40 * mm

        pdf.save()

    def gerar_a_partir_da_tabela(self):
        if self.df is None or not self.column_combo.get():
            messagebox.showwarning("Aviso", "Selecione um arquivo e uma coluna.")
            return

        codigos = [str(v) for v in self.df[self.column_combo.get()].dropna().tolist()]
        if not codigos:
            messagebox.showwarning("Aviso", "A coluna selecionada não possui dados válidos.")
            return

        formato = self.formato_saida.get()
        if formato == "pdf":
            caminho = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if caminho:
                self.gerar_pdf(codigos, caminho)
                messagebox.showinfo("Sucesso", "PDF gerado com sucesso.")
        elif formato == "zip":
            caminho = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
            if caminho:
                self.gerar_zip(codigos, caminho)
                messagebox.showinfo("Sucesso", "ZIP gerado com sucesso.")
        else:
            pasta = filedialog.askdirectory()
            if pasta:
                self.gerar_imagens(codigos, formato, pasta)
                messagebox.showinfo("Sucesso", f"Arquivos gerados em: {pasta}")


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()
