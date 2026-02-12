import io
import os
import queue
import tempfile
import threading
import traceback
import zipfile
from dataclasses import dataclass

import qrcode
from PIL import Image, ImageTk
from qrcode.image.svg import SvgImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService

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
        self.service = CodigoService()
        self.df = None
        self.arquivo_fonte = ""
        self.preview_image_ref = None
        self._preview_backend_error_shown = False

        self.qr_size = tk.IntVar(value=250)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.modo = tk.StringVar(value="texto")
        self.formato_saida = tk.StringVar(value="pdf")
        self.tipo_codigo = tk.StringVar(value="qrcode")

        self.prefixo_numerico = tk.StringVar(value="")
        self.sufixo_numerico = tk.StringVar(value="")
        self.max_codigos_por_lote = 5000
        self.max_tamanho_dado = 512

        self._criar_interface()
        self.atualizar_preview()
        self._geracao_em_andamento = False
        self._carregamento_em_andamento = False
        self.root.after(100, self.verificar_fila)

    def _criar_interface(self):
        topo = ttk.Frame(self.root, padding=10)
        topo.pack(fill="x")

        self.select_button = ttk.Button(topo, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.select_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

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

        self.progress_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        self.progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_label_var)
        self.progress_label.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="determinate", maximum=100)
        self.progress_bar.pack(fill="x", pady=(4, 0))
        self.progress_frame.pack_forget()

    def atualizar_controles_formato(self):
        if self.modo.get() == "texto":
            self.numerico_controls.grid_forget()
            self.texto_controls.grid(row=3, column=0, columnspan=3, sticky="w", padx=5, pady=3)
        else:
            self.texto_controls.grid_forget()
            self.numerico_controls.grid(row=3, column=0, columnspan=3, sticky="w", padx=5, pady=3)

    def selecionar_arquivo(self):
        if self._carregamento_em_andamento or self._geracao_em_andamento:
            return

        caminho = filedialog.askopenfilename(
            title="Selecione CSV ou Excel",
            filetypes=[("Arquivos de dados", "*.csv *.xlsx")],
        )
        if not caminho:
            return

        self._carregamento_em_andamento = True
        self.progress_bar.stop()
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar["value"] = 0
        self.progress_bar.start(10)
        self.progress_label_var.set("Carregando arquivo, aguarde...")
        self.progress_frame.pack(fill="x")
        self.select_button.configure(state="disabled")
        self.generate_button.configure(state="disabled")

        worker = threading.Thread(target=self._executar_carregamento, args=(caminho,), daemon=True)
        worker.start()

    def _executar_carregamento(self, caminho):
        try:
            tabela = self._carregar_tabela(caminho)
            self.fila.put({"tipo": "carregamento_sucesso", "caminho": caminho, "tabela": tabela})
        except Exception as exc:
            self.fila.put({"tipo": "carregamento_erro", "msg": str(exc)})

    def _formatar_excecao(self, exc: Exception, contexto: str) -> str:
        return self.service.formatar_excecao(exc, contexto)

    def _carregar_tabela(self, caminho):
        """Carrega CSV/XLSX com fallback quando pandas/numpy não estiverem disponíveis."""
        return self.service.carregar_tabela(caminho)

    def _obter_colunas(self, tabela):
        return self.service.obter_colunas(tabela)

    def _obter_valores_coluna(self, tabela, coluna):
        return self.service.obter_valores_coluna(tabela, coluna)

    def _build_config(self) -> GeracaoConfig:
        return GeracaoConfig(
            qr_size=int(self.qr_size.get()),
            foreground=self.qr_foreground_color.get(),
            background=self.qr_background_color.get(),
            tipo_codigo=self.tipo_codigo.get(),
            modo=self.modo.get(),
            prefixo=self.prefixo_numerico.get(),
            sufixo=self.sufixo_numerico.get(),
            max_codigos_por_lote=self.max_codigos_por_lote,
            max_tamanho_dado=self.max_tamanho_dado,
        )

    def _validar_parametros_geracao(self, codigos, cfg: GeracaoConfig | None = None):
        cfg = cfg or self._build_config()
        return self.service.validar_parametros_geracao(codigos, cfg)

    def _sanitizar_nome_arquivo(self, nome: str, fallback: str) -> str:
        return self.service.sanitizar_nome_arquivo(nome, fallback)

    def _normalizar_dado(self, valor: str, cfg: GeracaoConfig | None = None) -> str:
        cfg = cfg or self._build_config()
        return self.service.normalizar_dado(valor, cfg)

    def _gerar_imagem_obj(self, dado: str, cfg: GeracaoConfig | None = None) -> Image.Image:
        cfg = cfg or self._build_config()
        return self.service.gerar_imagem_obj(dado, cfg)

    def atualizar_preview(self):
        cfg = self._build_config()
        amostra = self._normalizar_dado("123456789", cfg)
        try:
            img = self._gerar_imagem_obj(amostra, cfg)
            img.thumbnail((260, 260))
            self.preview_image_ref = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self.preview_image_ref, text="")
            self._preview_backend_error_shown = False
        except RuntimeError as exc:
            # Evita quebrar callback do Tkinter quando backend opcional do reportlab não está disponível.
            self.preview_label.configure(image="", text="Preview indisponível para barcode neste ambiente")
            if not self._preview_backend_error_shown:
                messagebox.showwarning("Dependência opcional ausente", str(exc))
                self._preview_backend_error_shown = True

    def gerar_imagens(self, codigos, formato, destino, emitir_sucesso=True):
        try:
            cfg = self._build_config()
            os.makedirs(destino, exist_ok=True)
            total = len(codigos)

            nomes_usados = set()
            for i, codigo in enumerate(codigos, start=1):
                dado = self._normalizar_dado(codigo, cfg)
                nome_base = self._sanitizar_nome_arquivo(codigo, f"codigo_{i}")
                nome_arquivo = nome_base
                sufixo = 2
                while nome_arquivo in nomes_usados:
                    nome_arquivo = f"{nome_base}_{sufixo}"
                    sufixo += 1
                nomes_usados.add(nome_arquivo)

                if formato == "svg":
                    if cfg.tipo_codigo == "barcode":
                        raise ValueError("Exportação SVG para código de barras não suportada nesta versão.")
                    qr = qrcode.make(dado, image_factory=SvgImage)
                    caminho_saida = os.path.join(destino, f"{nome_arquivo}.svg")
                    with open(caminho_saida, "wb") as f:
                        qr.save(f)
                else:
                    imagem = self._gerar_imagem_obj(dado, cfg)
                    imagem.save(os.path.join(destino, f"{nome_arquivo}.png"), format="PNG")

                self.fila.put({"tipo": "progresso", "atual": i, "total": total, "codigo": codigo})

            if emitir_sucesso:
                self.fila.put({"tipo": "sucesso", "caminho": destino})
        except (OSError, ValueError) as exc:
            raise RuntimeError(self._formatar_excecao(exc, "Erro ao gerar imagens")) from exc

    def gerar_zip(self, codigos, caminho_zip):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                self.gerar_imagens(codigos, "png", tmpdir, emitir_sucesso=False)
                with zipfile.ZipFile(caminho_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nome in os.listdir(tmpdir):
                        if nome.lower().endswith('.png'):
                            zf.write(os.path.join(tmpdir, nome), arcname=nome)
            self.fila.put({"tipo": "sucesso", "caminho": caminho_zip})
        except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
            raise RuntimeError(self._formatar_excecao(exc, "Erro ao gerar ZIP")) from exc

    def gerar_pdf(self, codigos, caminho_pdf):
        try:
            cfg = self._build_config()
            pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
            largura_pagina, altura_pagina = A4

            x = 20 * mm
            y = altura_pagina - 40 * mm
            tamanho = 35 * mm
            margem = 10 * mm
            total = len(codigos)

            for i, codigo in enumerate(codigos, start=1):
                imagem = self._gerar_imagem_obj(self._normalizar_dado(codigo, cfg), cfg)
                buffer = io.BytesIO()
                imagem.save(buffer, format="PNG")
                buffer.seek(0)
                image_reader = ImageReader(buffer)

                pdf.drawImage(image_reader, x, y, width=tamanho, height=tamanho, preserveAspectRatio=True)
                self.fila.put({"tipo": "progresso", "atual": i, "total": total, "codigo": codigo})

                x += tamanho + margem
                if x + tamanho > largura_pagina - 20 * mm:
                    x = 20 * mm
                    y -= tamanho + margem

                if y < 20 * mm:
                    pdf.showPage()
                    x = 20 * mm
                    y = altura_pagina - 40 * mm

            pdf.save()
            self.fila.put({"tipo": "sucesso", "caminho": caminho_pdf})
        except (OSError, ValueError, RuntimeError) as exc:
            raise RuntimeError(self._formatar_excecao(exc, "Erro ao gerar PDF")) from exc

    def _iniciar_progresso(self, total):
        self._geracao_em_andamento = True
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate", maximum=max(total, 1))
        self.progress_bar["value"] = 0
        self.progress_label_var.set(f"Gerando 0/{total}...")
        self.progress_frame.pack(fill="x")
        self.generate_button.configure(state="disabled")
        self.select_button.configure(state="disabled")

    def _finalizar_progresso(self):
        self._geracao_em_andamento = False
        self._carregamento_em_andamento = False
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        self.select_button.configure(state="normal")
        if self.df is not None and self.column_combo.get():
            self.generate_button.configure(state="normal")

    def verificar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg["tipo"] == "progresso":
                    atual = msg.get("atual", 0)
                    total = msg.get("total", 1)
                    self.progress_bar.configure(maximum=max(total, 1))
                    self.progress_bar["value"] = atual
                    self.progress_label_var.set(f"Gerando {atual}/{total}: {msg.get('codigo', '')}")
                elif msg["tipo"] == "sucesso":
                    self._finalizar_progresso()
                    messagebox.showinfo("Sucesso", f"Arquivo(s) gerado(s) em: {msg.get('caminho', '')}")
                elif msg["tipo"] == "erro":
                    self._finalizar_progresso()
                    detalhe = msg.get("detalhe", "")
                    erro_msg = msg.get("msg", "Falha durante a geração.")
                    if detalhe:
                        erro_msg = f"{erro_msg}\n\nDetalhes técnicos:\n{detalhe}"
                    messagebox.showerror("Erro", erro_msg)
                elif msg["tipo"] == "carregamento_sucesso":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self._carregamento_em_andamento = False
                    self.select_button.configure(state="normal")
                    self.df = msg["tabela"]
                    self.arquivo_fonte = msg["caminho"]
                    colunas = self._obter_colunas(self.df)
                    self.column_combo.configure(values=colunas, state="readonly")
                    if colunas:
                        self.column_combo.set(colunas[0])
                        self.generate_button.configure(state="normal")
                    else:
                        self.generate_button.configure(state="disabled")
                elif msg["tipo"] == "carregamento_erro":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self._carregamento_em_andamento = False
                    self.select_button.configure(state="normal")
                    self.column_combo.configure(state="disabled", values=[])
                    self.generate_button.configure(state="disabled")
                    messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {msg.get('msg', '')}")
        except queue.Empty:
            pass

        self.root.after(100, self.verificar_fila)

    def _executar_geracao(self, codigos, formato, destino):
        try:
            if formato == "pdf":
                self.gerar_pdf(codigos, destino)
            elif formato == "zip":
                self.gerar_zip(codigos, destino)
            else:
                self.gerar_imagens(codigos, formato, destino)
        except Exception as exc:
            self.fila.put({"tipo": "erro", "msg": str(exc), "detalhe": traceback.format_exc(limit=3)})

    def gerar_a_partir_da_tabela(self):
        if self._geracao_em_andamento or self._carregamento_em_andamento:
            return
        if self.df is None or not self.column_combo.get():
            messagebox.showwarning("Aviso", "Selecione um arquivo e uma coluna.")
            return

        codigos = self._obter_valores_coluna(self.df, self.column_combo.get())
        if not codigos:
            messagebox.showwarning("Aviso", "A coluna selecionada não possui dados válidos.")
            return

        try:
            cfg = self._build_config()
            codigos, invalidos = self._validar_parametros_geracao(codigos, cfg)
        except ValueError as exc:
            messagebox.showwarning("Validação", str(exc))
            return

        if invalidos:
            messagebox.showwarning(
                "Validação",
                f"{invalidos} registro(s) foram ignorados por não atenderem aos limites de entrada.",
            )

        formato = self.formato_saida.get()
        destino = None
        if formato == "pdf":
            destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        elif formato == "zip":
            destino = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
        else:
            destino = filedialog.askdirectory()

        if not destino:
            return

        self._iniciar_progresso(len(codigos))
        worker = threading.Thread(target=self._executar_geracao, args=(codigos, formato, destino), daemon=True)
        worker.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()

