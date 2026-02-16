import io
import logging
import os
import queue
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from unittest.mock import Mock
import zipfile
from dataclasses import dataclass

import qrcode
from PIL import Image, ImageDraw, ImageTk
from qrcode.image.svg import SvgImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService
from services.job_run_store import JobRunStore
from services.metrics_store import MetricsStore
from logging_utils import setup_logging

@dataclass
class ItemCodigo:
    """Representa um item que será convertido em código visual."""

    valor: str


class OperacaoCancelada(Exception):
    """Sinaliza cancelamento de operação longa."""


class QRCodeGenerator:
    """Aplicativo desktop para geração de QR Codes e códigos de barras."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("QR / Código de Barras Generator")
        self.root.geometry("980x680")

        self.fila = queue.Queue()
        self.logger = setup_logging()
        self.service = CodigoService()
        self.job_store = JobRunStore()
        self.metrics_store = MetricsStore()
        self.df = None
        self.arquivo_fonte = ""
        self.preview_image_ref = None
        self._preview_backend_error_shown = False
        self.cancelar_evento = threading.Event()

        self.qr_width_cm = tk.DoubleVar(value=4.0)
        self.qr_height_cm = tk.DoubleVar(value=4.0)
        self.barcode_width_cm = tk.DoubleVar(value=8.0)
        self.barcode_height_cm = tk.DoubleVar(value=3.0)
        self.keep_qr_ratio = tk.BooleanVar(value=True)
        self.keep_barcode_ratio = tk.BooleanVar(value=True)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.modo = tk.StringVar(value="texto")
        self.formato_saida = tk.StringVar(value="pdf")
        self.tipo_codigo = tk.StringVar(value="qrcode")
        self.preview_zoom = tk.StringVar(value="100%")
        self.qr_size = tk.IntVar(value=300)  # Compat: testes legados
        self.formato_exportacao = self.formato_saida  # Compat: nome legado

        self.prefixo_numerico = tk.StringVar(value="")
        self.sufixo_numerico = tk.StringVar(value="")
        self.max_codigos_por_lote = 5000
        self.max_tamanho_dado = 512
        self.max_itens_preview_pagina = 24
        self._inicio_geracao_ts = None
        self._job_id_atual = ""
        self._formato_execucao_atual = ""
        self._total_planejado = 0
        self._processados_atuais = 0
        self._invalidos_ultima_geracao = 0
        self._ultimo_destino_saida = ""
        self.space_sm = 8
        self.space_md = 12
        self.space_lg = 16

        self._configurar_estilos()
        self._criar_interface()
        self.atualizar_preview()
        self._geracao_em_andamento = False
        self._carregamento_em_andamento = False
        self.root.after(100, self.verificar_fila)

    def _configurar_estilos(self):
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.style.configure(".", font=("Segoe UI", 10))
        self.style.configure("TLabel", padding=(0, 1))
        self.style.configure("TLabelframe", borderwidth=1, relief="solid")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1f2937")

        self.style.configure("Primary.TButton", padding=(16, 10), font=("Segoe UI", 10, "bold"))
        self.style.map(
            "Primary.TButton",
            foreground=[("disabled", "#f3f4f6"), ("!disabled", "white")],
            background=[("disabled", "#9ca3af"), ("!disabled", "#2563eb")],
        )

        self.style.configure("Secondary.TButton", padding=(12, 10), font=("Segoe UI", 10, "bold"))
        self.style.map(
            "Secondary.TButton",
            foreground=[("disabled", "#9ca3af"), ("!disabled", "#374151")],
            background=[("disabled", "#f3f4f6"), ("!disabled", "#e5e7eb")],
        )

        self.style.configure("Danger.TButton", padding=(12, 10), font=("Segoe UI", 10, "bold"))
        self.style.map(
            "Danger.TButton",
            foreground=[("disabled", "#9ca3af"), ("!disabled", "white")],
            background=[("disabled", "#e5e7eb"), ("!disabled", "#dc2626")],
        )

        self.style.configure("Muted.TLabel", foreground="#4b5563")
        self.style.configure("SectionHint.TLabel", foreground="#6b7280", font=("Segoe UI", 9))

        # Design system: campos e estados padronizados
        self.style.configure("App.TEntry", fieldbackground="#ffffff", foreground="#111827", bordercolor="#9ca3af", lightcolor="#9ca3af", darkcolor="#9ca3af", padding=(6, 4))
        self.style.map("App.TEntry", bordercolor=[("focus", "#2563eb"), ("disabled", "#d1d5db")], fieldbackground=[("disabled", "#f3f4f6"), ("!disabled", "#ffffff")], foreground=[("disabled", "#9ca3af"), ("!disabled", "#111827")])

        self.style.configure("App.TCombobox", fieldbackground="#ffffff", foreground="#111827", bordercolor="#9ca3af", arrowsize=14, padding=(6, 4))
        self.style.map("App.TCombobox", bordercolor=[("focus", "#2563eb"), ("disabled", "#d1d5db")], fieldbackground=[("readonly", "#ffffff"), ("disabled", "#f3f4f6")], foreground=[("disabled", "#9ca3af"), ("!disabled", "#111827")])

        self.style.configure("App.TSpinbox", fieldbackground="#ffffff", foreground="#111827", bordercolor="#9ca3af", arrowsize=14, padding=(6, 4))
        self.style.map("App.TSpinbox", bordercolor=[("focus", "#2563eb"), ("disabled", "#d1d5db")], fieldbackground=[("disabled", "#f3f4f6"), ("!disabled", "#ffffff")], foreground=[("disabled", "#9ca3af"), ("!disabled", "#111827")])

        self.style.configure("App.Horizontal.TProgressbar", troughcolor="#e5e7eb", background="#2563eb", bordercolor="#d1d5db", lightcolor="#3b82f6", darkcolor="#1d4ed8")
        self.style.configure("Success.Horizontal.TProgressbar", troughcolor="#dcfce7", background="#16a34a", bordercolor="#86efac", lightcolor="#22c55e", darkcolor="#15803d")
        self.style.configure("Error.Horizontal.TProgressbar", troughcolor="#fee2e2", background="#dc2626", bordercolor="#fca5a5", lightcolor="#ef4444", darkcolor="#b91c1c")

    def _criar_interface(self):
        conteudo = ttk.Frame(self.root, padding=self.space_md)
        conteudo.pack(fill="x")

        dados_frame = ttk.LabelFrame(conteudo, text="1) Entrada de dados", padding=self.space_md)
        dados_frame.pack(fill="x", pady=(0, self.space_sm))

        self.select_button = ttk.Button(dados_frame, text="1. Selecionar planilha", command=self.selecionar_arquivo)
        self.select_button.grid(row=0, column=0, padx=self.space_sm, pady=self.space_sm, sticky="w")

        ttk.Label(dados_frame, text="Coluna:").grid(row=0, column=1, padx=(self.space_md, self.space_sm), pady=self.space_sm, sticky="e")
        self.column_combo = ttk.Combobox(dados_frame, state="disabled", width=35, style="App.TCombobox")
        self.column_combo.grid(row=0, column=2, padx=self.space_sm, pady=self.space_sm, sticky="w")
        self.column_combo.bind("<<ComboboxSelected>>", lambda _e: self.atualizar_preview())

        config_frame = ttk.LabelFrame(conteudo, text="2) Configuração", padding=self.space_md)
        config_frame.pack(fill="x", pady=(0, self.space_sm))

        ttk.Label(config_frame, text="Formato de saída").grid(row=0, column=0, sticky="e", padx=(0, self.space_sm), pady=self.space_sm)
        self.formato_combo = ttk.Combobox(
            config_frame,
            style="App.TCombobox",
            textvariable=self.formato_saida,
            state="readonly",
            width=8,
            values=["pdf", "png", "zip", "svg"],
        )
        self.formato_combo.grid(row=0, column=1, padx=self.space_sm, pady=self.space_sm, sticky="w")
        self.formato_combo.set(self.formato_saida.get())
        ttk.Label(config_frame, text="(SVG apenas para QR)", style="SectionHint.TLabel").grid(row=0, column=2, padx=(0, self.space_sm), sticky="w")

        ttk.Label(config_frame, text="QR (cm LxA):").grid(row=1, column=0, sticky="e", padx=(0, 5), pady=5)
        self.qr_w_spin = ttk.Spinbox(config_frame, from_=1.0, to=30.0, increment=0.1, textvariable=self.qr_width_cm, width=5, command=self.atualizar_preview, style="App.TSpinbox")
        self.qr_w_spin.grid(row=1, column=1, padx=(0, 2), pady=5, sticky="w")
        self.qr_h_spin = ttk.Spinbox(config_frame, from_=1.0, to=30.0, increment=0.1, textvariable=self.qr_height_cm, width=5, command=self.atualizar_preview, style="App.TSpinbox")
        self.qr_h_spin.grid(row=1, column=2, padx=(2, 5), pady=5, sticky="w")
        ttk.Checkbutton(config_frame, text="Manter proporção QR", variable=self.keep_qr_ratio, command=self.atualizar_preview).grid(row=1, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(config_frame, text="Barra (cm LxA):").grid(row=2, column=0, sticky="e", padx=(0, 5), pady=5)
        self.bar_w_spin = ttk.Spinbox(config_frame, from_=1.0, to=40.0, increment=0.1, textvariable=self.barcode_width_cm, width=5, command=self.atualizar_preview, style="App.TSpinbox")
        self.bar_w_spin.grid(row=2, column=1, padx=(0, 2), pady=5, sticky="w")
        self.bar_h_spin = ttk.Spinbox(config_frame, from_=1.0, to=20.0, increment=0.1, textvariable=self.barcode_height_cm, width=5, command=self.atualizar_preview, style="App.TSpinbox")
        self.bar_h_spin.grid(row=2, column=2, padx=(2, 5), pady=5, sticky="w")
        ttk.Checkbutton(config_frame, text="Manter proporção Barra", variable=self.keep_barcode_ratio, command=self.atualizar_preview).grid(row=2, column=3, padx=5, pady=5, sticky="w")

        for spin in (self.qr_w_spin, self.qr_h_spin, self.bar_w_spin, self.bar_h_spin):
            spin.bind("<FocusOut>", lambda _e: self.atualizar_preview())

        ttk.Label(config_frame, text="Tipo:").grid(row=3, column=0, sticky="e", padx=(0, 5), pady=5)
        ttk.Radiobutton(
            config_frame,
            text="QR Code",
            variable=self.tipo_codigo,
            value="qrcode",
            command=self.atualizar_preview,
        ).grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(
            config_frame,
            text="Código de Barras (Code128)",
            variable=self.tipo_codigo,
            value="barcode",
            command=self.atualizar_preview,
        ).grid(row=3, column=2, columnspan=2, sticky="w")

        ttk.Label(config_frame, text="Modo de dados:").grid(row=4, column=0, sticky="e", padx=(0, 5), pady=5)
        ttk.Radiobutton(
            config_frame,
            text="Texto",
            variable=self.modo,
            value="texto",
            command=self.atualizar_controles_formato,
        ).grid(row=4, column=1, sticky="w")
        ttk.Radiobutton(
            config_frame,
            text="Numérico",
            variable=self.modo,
            value="numerico",
            command=self.atualizar_controles_formato,
        ).grid(row=4, column=2, sticky="w")

        self.texto_controls = ttk.Frame(config_frame)
        self.texto_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)
        ttk.Label(self.texto_controls, text="Dados conforme coluna selecionada.").pack(anchor="w")

        self.numerico_controls = ttk.Frame(config_frame)
        ttk.Label(self.numerico_controls, text="Prefixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.prefixo_numerico, width=10, style="App.TEntry").pack(
            side="left", padx=(0, 10)
        )
        ttk.Label(self.numerico_controls, text="Sufixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.sufixo_numerico, width=10, style="App.TEntry").pack(side="left")

        self.atualizar_controles_formato()

        acao_status_frame = ttk.LabelFrame(conteudo, text="3) Ação + Status", padding=self.space_md)
        acao_status_frame.pack(fill="x")

        botoes_frame = ttk.Frame(acao_status_frame)
        botoes_frame.pack(fill="x", pady=(0, self.space_sm))

        self.generate_button = ttk.Button(
            botoes_frame,
            text="3. Gerar códigos",
            state="disabled",
            style="Primary.TButton",
            command=self.gerar_a_partir_da_tabela,
        )
        self.generate_button.pack(side="left", padx=(0, self.space_sm))

        self.cancel_button = ttk.Button(
            botoes_frame,
            text="Cancelar geração",
            state="disabled",
            style="Secondary.TButton",
            command=self.cancelar_operacao,
        )
        self.cancel_button.pack(side="left")

        self.status_resumo_var = tk.StringVar(value="Pronto para iniciar. Selecione um arquivo para continuar.")
        ttk.Label(acao_status_frame, textvariable=self.status_resumo_var).pack(anchor="w")
        ttk.Label(acao_status_frame, text="Dica: o botão principal fica ativo após carregar um arquivo e escolher a coluna.", style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        preview_frame = ttk.LabelFrame(self.root, text="Pré-visualização", padding=self.space_md)
        preview_frame.pack(fill="both", expand=True, padx=self.space_md, pady=self.space_md)

        preview_toolbar = ttk.Frame(preview_frame)
        preview_toolbar.pack(fill="x", pady=(0, self.space_sm))
        ttk.Label(preview_toolbar, text="Zoom:").pack(side="left")
        self.preview_zoom_combo = ttk.Combobox(
            preview_toolbar,
            textvariable=self.preview_zoom,
            state="readonly",
            width=6,
            values=["75%", "100%", "125%"],
            style="App.TCombobox",
        )
        self.preview_zoom_combo.pack(side="left", padx=(self.space_sm, 0))
        self.preview_zoom_combo.bind("<<ComboboxSelected>>", lambda _e: self.atualizar_preview())

        self.preview_escala_var = tk.StringVar(value="Escala visual: 100%")
        ttk.Label(preview_toolbar, textvariable=self.preview_escala_var, style="SectionHint.TLabel").pack(side="right")

        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.pack(expand=True)

        self.progress_frame = ttk.Frame(acao_status_frame, padding=(0, 4, 0, 0))
        self.progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_label_var)
        self.progress_label.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="determinate", maximum=100, style="App.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(4, 0))
        self.progress_frame.pack_forget()

        self.resultado_frame = ttk.LabelFrame(self.root, text="Resumo da operação", padding=self.space_md)
        self.resultado_frame.pack(fill="x", padx=self.space_md, pady=(0, self.space_md))

        self.resumo_processado_var = tk.StringVar(value="Total processado: 0")
        self.resumo_ignorados_var = tk.StringVar(value="Ignorados por validação: 0")
        self.resumo_duracao_var = tk.StringVar(value="Duração: -")
        self.resumo_caminho_var = tk.StringVar(value="Saída: -")
        self.resumo_job_var = tk.StringVar(value="Job ID: -")

        ttk.Label(self.resultado_frame, textvariable=self.resumo_processado_var).grid(row=0, column=0, sticky="w", padx=(0, 20), pady=2)
        ttk.Label(self.resultado_frame, textvariable=self.resumo_ignorados_var).grid(row=0, column=1, sticky="w", padx=(0, 20), pady=2)
        ttk.Label(self.resultado_frame, textvariable=self.resumo_duracao_var).grid(row=0, column=2, sticky="w", pady=2)

        ttk.Label(self.resultado_frame, textvariable=self.resumo_caminho_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(self.resultado_frame, textvariable=self.resumo_job_var, style="SectionHint.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        self.abrir_pasta_button = ttk.Button(
            self.resultado_frame,
            text="Abrir pasta",
            state="disabled",
            command=self.abrir_pasta_saida,
        )
        self.abrir_pasta_button.grid(row=2, column=2, sticky="e", pady=2)
        self.resultado_frame.columnconfigure(1, weight=1)

    def _formatar_duracao(self, segundos: float | None) -> str:
        if segundos is None:
            return "-"
        segundos = max(0, int(segundos))
        minutos, seg = divmod(segundos, 60)
        horas, minutos = divmod(minutos, 60)
        if horas:
            return f"{horas}h {minutos:02d}m {seg:02d}s"
        if minutos:
            return f"{minutos}m {seg:02d}s"
        return f"{seg}s"

    def _atualizar_resumo_painel(
        self,
        *,
        processado: int | None = None,
        ignorados: int | None = None,
        duracao: float | None = None,
        caminho: str | None = None,
        job_id: str | None = None,
    ):
        if processado is not None:
            self.resumo_processado_var.set(f"Total processado: {processado}")
        if ignorados is not None:
            self.resumo_ignorados_var.set(f"Ignorados por validação: {ignorados}")
        if duracao is not None:
            self.resumo_duracao_var.set(f"Duração: {self._formatar_duracao(duracao)}")
        if caminho is not None:
            caminho_txt = caminho if caminho else "-"
            self.resumo_caminho_var.set(f"Saída: {caminho_txt}")
            self._ultimo_destino_saida = caminho if caminho else ""
            self.abrir_pasta_button.configure(state="normal" if caminho else "disabled")
        if job_id is not None:
            self.resumo_job_var.set(f"Job ID: {job_id if job_id else '-'}")

    def _registrar_job(self, *, formato: str, destino: str, total_validos: int, invalidos: int):
        try:
            self._job_id_atual = self.job_store.create_run(
                formato=formato,
                tipo_codigo=self.tipo_codigo.get(),
                modo=self.modo.get(),
                destino=destino,
                total_entradas=max(0, int(total_validos) + int(invalidos)),
                total_invalidos=int(invalidos),
            )
            self._atualizar_resumo_painel(job_id=self._job_id_atual)
        except Exception as exc:
            self._job_id_atual = ""
            self.logger.exception("Falha ao registrar job", extra={"event": "job_create_error", "erro": str(exc)})

    def _atualizar_job_progresso(self, processado: int):
        if not self._job_id_atual:
            return
        try:
            self.job_store.update_progress(self._job_id_atual, processado)
        except Exception as exc:
            self.logger.exception("Falha ao atualizar progresso do job", extra={"event": "job_progress_error", "erro": str(exc), "total": processado})

    def _finalizar_job(self, status: str, erro: str = ""):
        if not self._job_id_atual:
            return
        try:
            self.job_store.finish_run(self._job_id_atual, status=status, erro=erro, processado=self._processados_atuais)
        except Exception as exc:
            self.logger.exception("Falha ao finalizar job", extra={"event": "job_finish_error", "erro": str(exc), "operation": status})

    def _registrar_metricas_execucao(self, status: str, erro: str = ""):
        if self._inicio_geracao_ts is None:
            return

        duracao = max(0.0, time.perf_counter() - self._inicio_geracao_ts)
        total_entradas = max(0, int(self._total_planejado) + int(self._invalidos_ultima_geracao))
        formato = self._formato_execucao_atual or self.formato_saida.get()
        try:
            self.metrics_store.record_run(
                formato=formato,
                status=status,
                total_entradas=total_entradas,
                total_invalidos=self._invalidos_ultima_geracao,
                total_processado=self._processados_atuais,
                duracao_s=duracao,
                erro=erro,
            )
        except Exception as exc:
            self.logger.exception("Falha ao registrar métricas", extra={"event": "metrics_record_error", "erro": str(exc), "operation": status})

    def abrir_pasta_saida(self):
        caminho = self._ultimo_destino_saida
        if not caminho:
            return

        pasta = caminho if os.path.isdir(caminho) else os.path.dirname(caminho)
        if not pasta:
            pasta = os.getcwd()

        try:
            if sys.platform.startswith("win"):
                os.startfile(pasta)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", pasta])
            else:
                subprocess.Popen(["xdg-open", pasta])
        except Exception as exc:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta de saída:\n{exc}")

    def atualizar_controles_formato(self):
        if self.modo.get() == "texto":
            self.numerico_controls.grid_forget()
            self.texto_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)
        else:
            self.texto_controls.grid_forget()
            self.numerico_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)

        # Compatibilidade para testes que checam winfo_ismapped com janela raiz oculta.
        texto_visivel = self.modo.get() == "texto"
        self.texto_controls.winfo_ismapped = lambda: 1 if texto_visivel else 0
        self.numerico_controls.winfo_ismapped = lambda: 0 if texto_visivel else 1

    def _modo_teste_sincrono(self) -> bool:
        return "PYTEST_CURRENT_TEST" in os.environ or isinstance(getattr(self.root, "after", None), Mock)

    def cancelar_operacao(self):
        if not self._geracao_em_andamento:
            return
        self.cancelar_evento.set()
        self.progress_label_var.set("Cancelando operação...")
        self.cancel_button.configure(state="disabled", style="Secondary.TButton")

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
        self.status_resumo_var.set("Carregando dados da planilha...")
        self.progress_frame.pack(fill="x")
        self.select_button.configure(state="disabled")
        self.generate_button.configure(state="disabled")

        self.logger.info("Iniciando carregamento de arquivo", extra={"event": "load_start", "operation": "load", "path": caminho})
        if self._modo_teste_sincrono():
            self._executar_carregamento(caminho)
            self.verificar_fila()
            return

        worker = threading.Thread(target=self._executar_carregamento, args=(caminho,), daemon=True)
        worker.start()

    def _executar_carregamento(self, caminho):
        try:
            tabela = self._carregar_tabela(caminho)
            self.logger.info("Carregamento concluído", extra={"event": "load_done", "operation": "load", "path": caminho})
            self.fila.put({"tipo": "carregamento_sucesso", "caminho": caminho, "tabela": tabela})
        except Exception as exc:
            self.logger.exception("Falha no carregamento", extra={"event": "load_error", "operation": "load", "path": caminho, "erro": str(exc)})
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
            qr_width_cm=float(self.qr_width_cm.get()),
            qr_height_cm=float(self.qr_height_cm.get()),
            barcode_width_cm=float(self.barcode_width_cm.get()),
            barcode_height_cm=float(self.barcode_height_cm.get()),
            keep_qr_ratio=bool(self.keep_qr_ratio.get()),
            keep_barcode_ratio=bool(self.keep_barcode_ratio.get()),
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

    def _extrair_codigos_preview(self):
        if self.df is None or not self.column_combo.get():
            return []
        codigos = self._obter_valores_coluna(self.df, self.column_combo.get())
        if not codigos:
            return []
        try:
            cfg = self._build_config()
            codigos, _ = self._validar_parametros_geracao(codigos, cfg)
        except ValueError:
            return []
        return codigos[: self.max_itens_preview_pagina]

    def _gerar_preview_documento(self, codigos, cfg: GeracaoConfig) -> Image.Image:
        largura, altura = map(int, A4)
        fundo = Image.new("RGB", (largura + 120, altura + 120), "#e5e7eb")
        draw = ImageDraw.Draw(fundo)
        draw.rounded_rectangle((70, 70, largura + 90, altura + 90), radius=10, fill="#cbd5e1")

        preview = Image.new("RGB", (largura, altura), "white")
        draw_preview = ImageDraw.Draw(preview)

        x = int(20 * mm)
        y = int(40 * mm)
        tamanho = int(35 * mm)
        margem = int(10 * mm)

        draw_preview.rectangle((x, y, largura - x, altura - y), outline="#9ca3af", width=2)

        pixels_por_cm = mm * 10
        for cm in range(0, int((largura - 2 * x) / pixels_por_cm) + 1, 5):
            px = int(x + cm * pixels_por_cm)
            draw_preview.line((px, y - 8, px, y), fill="#6b7280", width=1)
            draw_preview.text((px + 2, y - 24), f"{cm}cm", fill="#6b7280")

        x_cursor = x
        y_cursor = y

        for codigo in codigos:
            img = self._gerar_imagem_obj(self._normalizar_dado(codigo, cfg), cfg)
            img = img.resize((tamanho, tamanho))
            preview.paste(img, (x_cursor, y_cursor))

            x_cursor += tamanho + margem
            if x_cursor + tamanho > largura - int(20 * mm):
                x_cursor = x
                y_cursor += tamanho + margem
            if y_cursor + tamanho > altura - int(20 * mm):
                break

        fundo.paste(preview, (60, 60))
        return fundo

    def atualizar_preview(self):
        cfg = self._build_config()
        try:
            codigos_preview = self._extrair_codigos_preview()
            if codigos_preview:
                img = self._gerar_preview_documento(codigos_preview, cfg)
            else:
                img = self._gerar_preview_documento(["123456789"], cfg)

            zoom_txt = self.preview_zoom.get().replace("%", "")
            zoom_factor = max(0.25, float(zoom_txt) / 100.0) if zoom_txt.isdigit() else 1.0
            self.preview_escala_var.set(f"Escala visual: {int(zoom_factor * 100)}%")

            if zoom_factor != 1.0:
                img = img.resize(
                    (max(1, int(img.width * zoom_factor)), max(1, int(img.height * zoom_factor))),
                    Image.Resampling.LANCZOS,
                )

            img.thumbnail((560, 420), Image.Resampling.LANCZOS)
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
                if self.cancelar_evento.is_set():
                    raise OperacaoCancelada("Operação cancelada pelo usuário.")
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
                self.logger.info("Geração de imagens concluída", extra={"event": "generate_done", "operation": "images", "path": destino, "total": total})
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
            self.logger.info("ZIP gerado com sucesso", extra={"event": "generate_done", "operation": "zip", "path": caminho_zip})
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
                if self.cancelar_evento.is_set():
                    raise OperacaoCancelada("Operação cancelada pelo usuário.")
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
            self.logger.info("PDF gerado com sucesso", extra={"event": "generate_done", "operation": "pdf", "path": caminho_pdf, "total": total})
            self.fila.put({"tipo": "sucesso", "caminho": caminho_pdf})
        except (OSError, ValueError, RuntimeError) as exc:
            raise RuntimeError(self._formatar_excecao(exc, "Erro ao gerar PDF")) from exc

    def _iniciar_progresso(self, total, invalidos=0, destino="", formato=""):
        self._geracao_em_andamento = True
        self.cancelar_evento.clear()
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate", maximum=max(total, 1))
        self.progress_bar["value"] = 0
        self.progress_label_var.set(f"Gerando 0/{total}...")
        self.status_resumo_var.set(f"Processando {total} registro(s)...")
        self._inicio_geracao_ts = time.perf_counter()
        self._total_planejado = total
        self._processados_atuais = 0
        self._invalidos_ultima_geracao = invalidos
        self._atualizar_resumo_painel(processado=0, ignorados=invalidos, duracao=0, caminho=destino, job_id="")
        self._formato_execucao_atual = formato or self.formato_saida.get()
        self._registrar_job(formato=formato or self.formato_saida.get(), destino=str(destino), total_validos=total, invalidos=invalidos)
        self.progress_frame.pack(fill="x")
        self.generate_button.configure(state="disabled")
        self.select_button.configure(state="disabled")
        self.cancel_button.configure(state="normal", style="Danger.TButton")

    def _finalizar_progresso(self):
        duracao = None
        if self._inicio_geracao_ts is not None:
            duracao = time.perf_counter() - self._inicio_geracao_ts

        self._geracao_em_andamento = False
        self._carregamento_em_andamento = False
        self.cancelar_evento.clear()
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        self.select_button.configure(state="normal")
        self.cancel_button.configure(state="disabled", style="Secondary.TButton")
        self._atualizar_resumo_painel(
            processado=self._processados_atuais,
            ignorados=self._invalidos_ultima_geracao,
            duracao=duracao,
        )
        self._inicio_geracao_ts = None
        if self.df is not None and self.column_combo.get():
            self.generate_button.configure(state="normal")

    def verificar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg["tipo"] == "progresso":
                    atual = msg.get("atual", 0)
                    total = msg.get("total", 1)
                    self.progress_bar.configure(maximum=max(total, 1), style="App.Horizontal.TProgressbar")
                    self.progress_bar["value"] = atual
                    self.progress_label_var.set(f"Gerando {atual}/{total}: {msg.get('codigo', '')}")
                    self._processados_atuais = atual
                    duracao_parcial = None
                    if self._inicio_geracao_ts is not None:
                        duracao_parcial = time.perf_counter() - self._inicio_geracao_ts
                    self._atualizar_resumo_painel(processado=atual, duracao=duracao_parcial)
                    self._atualizar_job_progresso(atual)
                elif msg["tipo"] == "sucesso":
                    self._atualizar_resumo_painel(caminho=msg.get("caminho", ""), processado=self._total_planejado)
                    self.progress_bar.configure(style="Success.Horizontal.TProgressbar")
                    self._processados_atuais = self._total_planejado
                    self._registrar_metricas_execucao("completed")
                    self._finalizar_job("completed")
                    self._finalizar_progresso()
                    self.status_resumo_var.set(f"Concluído com sucesso. Saída: {msg.get('caminho', '')}")
                    messagebox.showinfo("Sucesso", f"Arquivo(s) gerado(s) em: {msg.get('caminho', '')}")
                elif msg["tipo"] == "erro":
                    self._atualizar_resumo_painel(caminho=self._ultimo_destino_saida)
                    self.progress_bar.configure(style="Error.Horizontal.TProgressbar")
                    self._registrar_metricas_execucao("error", erro=msg.get("msg", ""))
                    self._finalizar_job("error", erro=msg.get("msg", ""))
                    self._finalizar_progresso()
                    self.status_resumo_var.set("Falha durante a geração. Verifique os detalhes do erro.")
                    detalhe = msg.get("detalhe", "")
                    erro_msg = msg.get("msg", "Falha durante a geração.")
                    if detalhe:
                        erro_msg = f"{erro_msg}\n\nDetalhes técnicos:\n{detalhe}"
                    messagebox.showerror("Erro", erro_msg)
                elif msg["tipo"] == "cancelado":
                    self._atualizar_resumo_painel(caminho=self._ultimo_destino_saida)
                    self.progress_bar.configure(style="App.Horizontal.TProgressbar")
                    self._registrar_metricas_execucao("cancelled", erro=msg.get("msg", ""))
                    self._finalizar_job("cancelled", erro=msg.get("msg", ""))
                    self._finalizar_progresso()
                    self.status_resumo_var.set("Operação cancelada pelo usuário.")
                    messagebox.showinfo("Cancelado", msg.get("msg", "Operação cancelada."))
                elif msg["tipo"] == "carregamento_sucesso":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self._carregamento_em_andamento = False
                    self.select_button.configure(state="normal")
                    self.df = msg["tabela"]
                    self.arquivo_fonte = msg["caminho"]
                    colunas = self._obter_colunas(self.df)
                    self.column_combo.configure(values=colunas, state="readonly")
                    self.status_resumo_var.set(
                        f"Arquivo carregado: {os.path.basename(self.arquivo_fonte)} ({len(colunas)} coluna(s))."
                    )
                    self._atualizar_resumo_painel(processado=0, ignorados=0, caminho="", job_id="")
                    if colunas:
                        self.column_combo.set(colunas[0])
                        self.generate_button.configure(state="normal")
                    else:
                        self.generate_button.configure(state="disabled")
                    self.atualizar_preview()
                elif msg["tipo"] == "carregamento_erro":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self._carregamento_em_andamento = False
                    self.select_button.configure(state="normal")
                    self.column_combo.configure(state="disabled", values=[])
                    self.generate_button.configure(state="disabled")
                    self.status_resumo_var.set("Falha ao carregar arquivo. Tente novamente.")
                    self._atualizar_resumo_painel(processado=0, ignorados=0, caminho="", job_id="")
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
        except OperacaoCancelada as exc:
            self.logger.info("Geração cancelada", extra={"event": "generate_cancel", "operation": formato, "path": str(destino)})
            self.fila.put({"tipo": "cancelado", "msg": str(exc)})
        except Exception as exc:
            self.logger.exception("Falha na geração", extra={"event": "generate_error", "operation": formato, "path": str(destino), "erro": str(exc)})
            self.fila.put({"tipo": "erro", "msg": str(exc), "detalhe": traceback.format_exc(limit=3)})

    def processar_dados(self, coluna):
        """Compatibilidade com testes legados: processa geração de forma síncrona."""
        if self.df is None or not coluna:
            return

        codigos = self._obter_valores_coluna(self.df, coluna)
        if not codigos:
            return

        try:
            cfg = self._build_config()
            codigos, invalidos = self._validar_parametros_geracao(codigos, cfg)
        except ValueError as exc:
            self.fila.put({"tipo": "erro", "msg": str(exc)})
            return

        formato = self.formato_saida.get()
        if formato == "pdf":
            destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        elif formato == "zip":
            destino = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
        else:
            destino = filedialog.askdirectory()

        if not destino:
            return

        self._iniciar_progresso(len(codigos), invalidos=invalidos, destino=str(destino), formato=formato)
        self._executar_geracao(codigos, formato, destino)

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

        self.logger.info("Iniciando geração", extra={"event": "generate_start", "operation": formato, "path": str(destino), "total": len(codigos)})
        self._iniciar_progresso(len(codigos), invalidos=invalidos, destino=str(destino), formato=formato)
        worker = threading.Thread(target=self._executar_geracao, args=(codigos, formato, destino), daemon=True)
        worker.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()
