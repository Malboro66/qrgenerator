import io
import json
import logging
import os
import queue
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import zipfile
from dataclasses import dataclass
from enum import Enum, auto

import qrcode
from PIL import Image, ImageDraw, ImageTk
from qrcode.image.svg import SvgImage
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_controller import AppController
from models.geracao_config import GeracaoConfig

# Equivalentes do ReportLab para evitar dependência em tempo de import.
MM_TO_POINTS = 72 / 25.4
mm = MM_TO_POINTS
A4 = (210 * mm, 297 * mm)


def _obter_modulos_pdf_reportlab():
    try:
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas as pdf_canvas
        return pdf_canvas, ImageReader
    except ImportError as exc:
        raise RuntimeError(
            "Exportação PDF requer a dependência opcional 'reportlab'. "
            "Instale com: pip install reportlab"
        ) from exc


@dataclass
class ItemCodigo:
    """Representa um item que será convertido em código visual."""

    valor: str


@dataclass
class InterfaceSections:
    dados_frame: ttk.LabelFrame | None = None
    config_frame: ttk.LabelFrame | None = None
    acao_status_frame: ttk.LabelFrame | None = None
    preview_frame: ttk.LabelFrame | None = None
    resultado_frame: ttk.LabelFrame | None = None


class OperacaoCancelada(Exception):
    """Sinaliza cancelamento de operação longa."""


class EstadoAplicacao(Enum):
    IDLE = auto()
    LOADING = auto()
    READY = auto()
    GENERATING = auto()
    CANCELLING = auto()
    ERROR = auto()


class QRCodeGenerator:
    """Aplicativo desktop para geração de QR Codes e códigos de barras."""

    def __init__(self, root: tk.Tk, controller: AppController | None = None):
        self.root = root
        self.root.title("")
        self.root.geometry("980x680")

        self.fila = queue.Queue()
        self.controller = controller or AppController.build_default()
        self.logger = self.controller.logger
        self.job_store = self.controller.job_store
        self.metrics_store = self.controller.metrics_store
        self.df = None
        self.arquivo_fonte = ""
        self.preview_image_ref = None
        self._preview_backend_error_shown = False
        self._preview_after_id = None
        self.preview_debounce_ms = 250
        self.pdf_export_disponivel = False
        self.barcode_disponivel = False
        self.motivos_dependencias_indisponiveis = []
        self.cancelar_evento = threading.Event()
        self.sections = InterfaceSections()

        self.qr_width_cm = tk.StringVar(value="4.0")
        self.qr_height_cm = tk.StringVar(value="4.0")
        self.barcode_width_cm = tk.StringVar(value="8.0")
        self.barcode_height_cm = tk.StringVar(value="3.0")
        self.keep_qr_ratio = tk.BooleanVar(value=True)
        self.keep_barcode_ratio = tk.BooleanVar(value=True)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.modo = tk.StringVar(value="texto")
        self.formato_saida = tk.StringVar(value="pdf")
        self.tipo_codigo = tk.StringVar(value="qrcode")
        self.barcode_model = tk.StringVar(value="code128")
        self.preview_zoom = tk.StringVar(value="100%")
        self.preview_preset = tk.StringVar(value="A4")
        self.preview_margin_cm = tk.StringVar(value="2.0")
        self.preview_spacing_cm = tk.StringVar(value="1.0")
        self.impressora_var = tk.StringVar(value="")
        self.copias_impressao = tk.IntVar(value=1)
        self.impressora_status_var = tk.StringVar(value="")

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
        self._arquivos_temporarios_impressao = []
        self.space_sm = 8
        self.space_md = 12
        self.space_lg = 16
        self.etapa_atual = 1
        self.estado_atual = EstadoAplicacao.IDLE

        self._configurar_estilos()
        self._verificar_dependencias_essenciais()
        self.root.title(self._t("app.title", "QR / Código de Barras Generator"))
        self.barcode_model_options = self.controller.obter_modelos_barcode()
        self.barcode_label_to_key = {rotulo: chave for chave, rotulo in self.barcode_model_options}
        self.barcode_key_to_label = {chave: rotulo for chave, rotulo in self.barcode_model_options}
        self._criar_interface()
        self.atualizar_preview()
        self._aplicar_estado_ui()
        self.root.after(100, self.verificar_fila)

    def _t(self, key: str, default: str = "", **kwargs) -> str:
        return self.controller.t(key, default, **kwargs)

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

        self.style.configure("Step.TButton", padding=(12, 6), font=("Segoe UI", 10, "bold"))
        self.style.configure("StepActive.TButton", padding=(12, 6), font=("Segoe UI", 10, "bold"))
        self.style.map(
            "StepActive.TButton",
            foreground=[("!disabled", "white")],
            background=[("!disabled", "#2563eb")],
        )

    def _verificar_dependencias_essenciais(self):
        self.motivos_dependencias_indisponiveis = []

        try:
            from reportlab.pdfgen import canvas as _pdf_canvas  # noqa: F401

            self.pdf_export_disponivel = True
        except Exception:
            self.pdf_export_disponivel = False
            self.motivos_dependencias_indisponiveis.append("PDF/Impressão indisponível (faltando reportlab).")

        try:
            from reportlab.graphics import renderPM as _render_pm  # noqa: F401
            from reportlab.graphics.barcode import createBarcodeDrawing as _barcode_factory  # noqa: F401

            self.barcode_disponivel = True
        except Exception:
            self.barcode_disponivel = False
            self.motivos_dependencias_indisponiveis.append("Código de barras indisponível (faltando backend renderPM).")

    def _criar_interface(self):
        conteudo = ttk.Frame(self.root, padding=self.space_md)
        conteudo.pack(fill="x")

        self.stepper_frame = ttk.Frame(conteudo)
        self.stepper_frame.pack(fill="x", pady=(0, self.space_sm))
        self.step1_button = ttk.Button(self.stepper_frame, text=self._t("step.input", "1. Entrada"), style="Step.TButton", command=lambda: self._definir_etapa(1))
        self.step1_button.pack(side="left", padx=(0, self.space_sm))
        self.step2_button = ttk.Button(self.stepper_frame, text=self._t("step.config", "2. Configuração"), style="Step.TButton", command=lambda: self._definir_etapa(2))
        self.step2_button.pack(side="left", padx=(0, self.space_sm))
        self.step3_button = ttk.Button(self.stepper_frame, text=self._t("step.action", "3. Ação"), style="Step.TButton", command=lambda: self._definir_etapa(3))
        self.step3_button.pack(side="left")

        self.dados_frame = ttk.LabelFrame(conteudo, text=self._t("section.input", "1) Entrada de dados"), padding=self.space_md)
        self.dados_frame.pack(fill="x", pady=(0, self.space_sm))
        self.sections.dados_frame = self.dados_frame

        self.select_button = ttk.Button(self.dados_frame, text=self._t("button.select_spreadsheet", "1. Selecionar planilha"), command=self.selecionar_arquivo)
        self.select_button.grid(row=0, column=0, padx=self.space_sm, pady=self.space_sm, sticky="w")

        ttk.Label(self.dados_frame, text=self._t("label.column", "Coluna:")).grid(row=0, column=1, padx=(self.space_md, self.space_sm), pady=self.space_sm, sticky="e")
        self.column_combo = ttk.Combobox(self.dados_frame, state="disabled", width=35, style="App.TCombobox")
        self.column_combo.grid(row=0, column=2, padx=self.space_sm, pady=self.space_sm, sticky="w")
        self.column_combo.bind("<<ComboboxSelected>>", self._ao_selecionar_coluna)

        self.config_frame = ttk.LabelFrame(conteudo, text=self._t("section.config", "2) Configuração"), padding=self.space_md)
        self.config_frame.pack(fill="x", pady=(0, self.space_sm))
        self.sections.config_frame = self.config_frame

        ttk.Label(self.config_frame, text=self._t("label.output_format", "Formato de saída")).grid(row=0, column=0, sticky="e", padx=(0, self.space_sm), pady=self.space_sm)
        self.formato_combo = ttk.Combobox(
            self.config_frame,
            style="App.TCombobox",
            textvariable=self.formato_saida,
            state="readonly",
            width=8,
            values=["pdf", "png", "zip", "svg"],
        )
        self.formato_combo.grid(row=0, column=1, padx=self.space_sm, pady=self.space_sm, sticky="w")
        self.formato_combo.set(self.formato_saida.get())
        self.formato_combo.bind("<<ComboboxSelected>>", self._ao_alterar_formato_saida)
        ttk.Label(self.config_frame, text=self._t("hint.svg_only_qr", "(SVG apenas para QR)"), style="SectionHint.TLabel").grid(row=0, column=2, padx=(0, self.space_sm), sticky="w")
        formatos_disponiveis = ["png", "zip", "svg"]
        if self.pdf_export_disponivel:
            formatos_disponiveis = ["pdf", "png", "zip", "svg", "imprimir"]
        self.formato_combo.configure(values=formatos_disponiveis)
        if self.formato_saida.get() not in formatos_disponiveis:
            self.formato_saida.set("png")
            self.formato_combo.set("png")

        ttk.Label(self.config_frame, text="QR (cm LxA):").grid(row=1, column=0, sticky="e", padx=(0, 5), pady=5)
        self.qr_w_spin = ttk.Spinbox(self.config_frame, from_=1.0, to=30.0, increment=0.1, textvariable=self.qr_width_cm, width=5, command=self.solicitar_atualizacao_preview, style="App.TSpinbox")
        self.qr_w_spin.grid(row=1, column=1, padx=(0, 2), pady=5, sticky="w")
        self.qr_h_spin = ttk.Spinbox(self.config_frame, from_=1.0, to=30.0, increment=0.1, textvariable=self.qr_height_cm, width=5, command=self.solicitar_atualizacao_preview, style="App.TSpinbox")
        self.qr_h_spin.grid(row=1, column=2, padx=(2, 5), pady=5, sticky="w")
        ttk.Checkbutton(self.config_frame, text="Manter proporção QR", variable=self.keep_qr_ratio, command=self.solicitar_atualizacao_preview).grid(row=1, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(self.config_frame, text="Barra (cm LxA):").grid(row=2, column=0, sticky="e", padx=(0, 5), pady=5)
        self.bar_w_spin = ttk.Spinbox(self.config_frame, from_=1.0, to=40.0, increment=0.1, textvariable=self.barcode_width_cm, width=5, command=self.solicitar_atualizacao_preview, style="App.TSpinbox")
        self.bar_w_spin.grid(row=2, column=1, padx=(0, 2), pady=5, sticky="w")
        self.bar_h_spin = ttk.Spinbox(self.config_frame, from_=1.0, to=20.0, increment=0.1, textvariable=self.barcode_height_cm, width=5, command=self.solicitar_atualizacao_preview, style="App.TSpinbox")
        self.bar_h_spin.grid(row=2, column=2, padx=(2, 5), pady=5, sticky="w")
        ttk.Checkbutton(self.config_frame, text="Manter proporção Barra", variable=self.keep_barcode_ratio, command=self.solicitar_atualizacao_preview).grid(row=2, column=3, padx=5, pady=5, sticky="w")

        for spin in (self.qr_w_spin, self.qr_h_spin, self.bar_w_spin, self.bar_h_spin):
            spin.bind("<FocusOut>", lambda _e: self.solicitar_atualizacao_preview())
            spin.bind("<KeyRelease>", lambda _e: self.solicitar_atualizacao_preview())

        ttk.Label(self.config_frame, text=self._t("label.type", "Tipo:")).grid(row=3, column=0, sticky="e", padx=(0, 5), pady=5)
        self.tipo_qr_radio = ttk.Radiobutton(
            self.config_frame,
            text=self._t("type.qr", "QR Code"),
            variable=self.tipo_codigo,
            value="qrcode",
            command=self._ao_alterar_tipo_codigo,
        )
        self.tipo_qr_radio.grid(row=3, column=1, sticky="w")
        self.tipo_barcode_radio = ttk.Radiobutton(
            self.config_frame,
            text=self._t("type.barcode", "Código de Barras"),
            variable=self.tipo_codigo,
            value="barcode",
            command=self._ao_alterar_tipo_codigo,
        )
        self.tipo_barcode_radio.grid(row=3, column=2, sticky="w")
        ttk.Label(self.config_frame, text=self._t("label.barcode_model", "Modelo de etiqueta:")).grid(row=3, column=3, sticky="e", padx=(0, 5), pady=5)
        self.barcode_model_combo = ttk.Combobox(
            self.config_frame,
            state="readonly",
            width=30,
            style="App.TCombobox",
            values=[rotulo for _chave, rotulo in self.barcode_model_options],
        )
        self.barcode_model_combo.grid(row=3, column=4, padx=(2, 5), pady=5, sticky="w")
        self.barcode_model_combo.set(self.barcode_key_to_label.get(self.barcode_model.get(), "Código 128"))
        self.barcode_model_combo.bind("<<ComboboxSelected>>", self._ao_alterar_modelo_barcode)
        aviso_dependencias = "Todos os recursos disponíveis."
        if self.motivos_dependencias_indisponiveis:
            aviso_dependencias = " | ".join(self.motivos_dependencias_indisponiveis)
        self.dependency_status_var = tk.StringVar(value=aviso_dependencias)
        ttk.Label(self.config_frame, textvariable=self.dependency_status_var, style="Muted.TLabel").grid(
            row=7, column=0, columnspan=5, sticky="w", padx=5, pady=(2, 0)
        )

        ttk.Label(self.config_frame, text=self._t("label.data_mode", "Modo de dados:")).grid(row=4, column=0, sticky="e", padx=(0, 5), pady=5)
        ttk.Radiobutton(
            self.config_frame,
            text=self._t("mode.text", "Texto"),
            variable=self.modo,
            value="texto",
            command=self.atualizar_controles_formato,
        ).grid(row=4, column=1, sticky="w")
        ttk.Radiobutton(
            self.config_frame,
            text=self._t("mode.numeric", "Numérico"),
            variable=self.modo,
            value="numerico",
            command=self.atualizar_controles_formato,
        ).grid(row=4, column=2, sticky="w")

        self.texto_controls = ttk.Frame(self.config_frame)
        self.texto_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)
        ttk.Label(self.texto_controls, text="Dados conforme coluna selecionada.").pack(anchor="w")

        self.numerico_controls = ttk.Frame(self.config_frame)
        ttk.Label(self.numerico_controls, text="Prefixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.prefixo_numerico, width=10, style="App.TEntry").pack(
            side="left", padx=(0, 10)
        )
        ttk.Label(self.numerico_controls, text="Sufixo:").pack(side="left", padx=(0, 5))
        ttk.Entry(self.numerico_controls, textvariable=self.sufixo_numerico, width=10, style="App.TEntry").pack(side="left")

        self.atualizar_controles_formato()
        self._atualizar_controles_tipo_codigo()
        self._aplicar_disponibilidade_dependencias()

        self.impressao_frame = ttk.LabelFrame(self.config_frame, text="Configuração de impressão", padding=self.space_sm)
        self.impressao_frame.grid(row=6, column=0, columnspan=4, sticky="ew", padx=5, pady=(self.space_sm, 0))
        ttk.Label(self.impressao_frame, text="Impressora:").grid(row=0, column=0, sticky="e", padx=(0, self.space_sm), pady=4)
        self.impressora_combo = ttk.Combobox(
            self.impressao_frame,
            textvariable=self.impressora_var,
            state="readonly",
            width=38,
            style="App.TCombobox",
            values=[],
        )
        self.impressora_combo.grid(row=0, column=1, sticky="w", pady=4)
        ttk.Button(
            self.impressao_frame,
            text="Atualizar",
            style="Secondary.TButton",
            command=lambda: self.atualizar_lista_impressoras(notificar=True),
        ).grid(row=0, column=2, padx=(self.space_sm, 0), pady=4, sticky="w")

        ttk.Label(self.impressao_frame, text="Cópias:").grid(row=1, column=0, sticky="e", padx=(0, self.space_sm), pady=4)
        self.copias_spin = ttk.Spinbox(
            self.impressao_frame,
            from_=1,
            to=50,
            increment=1,
            textvariable=self.copias_impressao,
            width=6,
            style="App.TSpinbox",
        )
        self.copias_spin.grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(self.impressao_frame, textvariable=self.impressora_status_var, style="SectionHint.TLabel").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(2, 0)
        )
        self.impressao_frame.columnconfigure(1, weight=1)
        self.atualizar_lista_impressoras()
        self._atualizar_controles_impressao()

        self.acao_status_frame = ttk.LabelFrame(conteudo, text=self._t("section.action_status", "3) Ação + Status"), padding=self.space_md)
        self.acao_status_frame.pack(fill="x")
        self.sections.acao_status_frame = self.acao_status_frame

        botoes_frame = ttk.Frame(self.acao_status_frame)
        botoes_frame.pack(fill="x", pady=(0, self.space_sm))

        self.generate_button = ttk.Button(
            botoes_frame,
            text=self._t("button.generate", "3. Gerar códigos"),
            state="disabled",
            style="Primary.TButton",
            command=self.gerar_a_partir_da_tabela,
        )
        self.generate_button.pack(side="left", padx=(0, self.space_sm))

        self.test_print_button = ttk.Button(
            botoes_frame,
            text="Imprimir teste (1 código)",
            state="disabled",
            style="Secondary.TButton",
            command=self.imprimir_teste,
        )
        self.test_print_button.pack(side="left", padx=(0, self.space_sm))

        self.cancel_button = ttk.Button(
            botoes_frame,
            text=self._t("button.cancel", "Cancelar geração"),
            state="disabled",
            style="Secondary.TButton",
            command=self.cancelar_operacao,
        )
        self.cancel_button.pack(side="left")

        self.status_resumo_var = tk.StringVar(value=self._t("status.ready", "Pronto para iniciar. Selecione um arquivo para continuar."))
        ttk.Label(self.acao_status_frame, textvariable=self.status_resumo_var).pack(anchor="w")
        ttk.Label(self.acao_status_frame, text="Dica: o botão principal fica ativo após carregar um arquivo e escolher a coluna.", style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        preview_frame = ttk.LabelFrame(self.root, text="Pré-visualização", padding=self.space_md)
        preview_frame.pack(fill="both", expand=True, padx=self.space_md, pady=self.space_md)
        self.sections.preview_frame = preview_frame

        preview_toolbar = ttk.Frame(preview_frame)
        preview_toolbar.pack(fill="x", pady=(0, self.space_sm))
        ttk.Label(preview_toolbar, text="Preset:").pack(side="left")
        self.preview_preset_combo = ttk.Combobox(
            preview_toolbar,
            textvariable=self.preview_preset,
            state="readonly",
            width=20,
            values=["A4", "Etiqueta 8x10.5 cm", "Etiqueta 60x40 mm"],
            style="App.TCombobox",
        )
        self.preview_preset_combo.pack(side="left", padx=(self.space_sm, self.space_md))
        self.preview_preset_combo.bind("<<ComboboxSelected>>", lambda _e: self.solicitar_atualizacao_preview())

        ttk.Label(preview_toolbar, text="Margem (cm):").pack(side="left")
        self.preview_margin_spin = ttk.Spinbox(
            preview_toolbar,
            from_=0.2,
            to=5.0,
            increment=0.1,
            textvariable=self.preview_margin_cm,
            width=5,
            command=self.solicitar_atualizacao_preview,
            style="App.TSpinbox",
        )
        self.preview_margin_spin.pack(side="left", padx=(self.space_sm, self.space_md))

        ttk.Label(preview_toolbar, text="Espaço (cm):").pack(side="left")
        self.preview_spacing_spin = ttk.Spinbox(
            preview_toolbar,
            from_=0.1,
            to=5.0,
            increment=0.1,
            textvariable=self.preview_spacing_cm,
            width=5,
            command=self.solicitar_atualizacao_preview,
            style="App.TSpinbox",
        )
        self.preview_spacing_spin.pack(side="left", padx=(self.space_sm, self.space_md))

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
        self.preview_zoom_combo.bind("<<ComboboxSelected>>", lambda _e: self.solicitar_atualizacao_preview())
        self.preview_margin_spin.bind("<FocusOut>", lambda _e: self.solicitar_atualizacao_preview())
        self.preview_spacing_spin.bind("<FocusOut>", lambda _e: self.solicitar_atualizacao_preview())
        self.preview_margin_spin.bind("<KeyRelease>", lambda _e: self.solicitar_atualizacao_preview())
        self.preview_spacing_spin.bind("<KeyRelease>", lambda _e: self.solicitar_atualizacao_preview())

        self.preview_escala_var = tk.StringVar(value="Escala visual: 100%")
        ttk.Label(preview_toolbar, textvariable=self.preview_escala_var, style="SectionHint.TLabel").pack(side="right")

        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.pack(expand=True)

        self.progress_frame = ttk.Frame(self.acao_status_frame, padding=(0, 4, 0, 0))
        self.progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_label_var)
        self.progress_label.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="determinate", maximum=100, style="App.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(4, 0))
        self.progress_frame.pack_forget()

        self.resultado_frame = ttk.LabelFrame(self.root, text="Resumo da operação", padding=self.space_md)
        self.resultado_frame.pack(fill="x", padx=self.space_md, pady=(0, self.space_md))
        self.sections.resultado_frame = self.resultado_frame

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
        self._definir_etapa(1, forcar=True)

    def _atualizar_stepper_visual(self):
        botoes = {
            1: self.step1_button,
            2: self.step2_button,
            3: self.step3_button,
        }
        for etapa, botao in botoes.items():
            botao.configure(style="StepActive.TButton" if etapa == self.etapa_atual else "Step.TButton")

        self.step2_button.configure(state="normal" if self.df is not None else "disabled")
        self.step3_button.configure(state="normal" if (self.df is not None and self.column_combo.get()) else "disabled")

    def _definir_etapa(self, etapa: int, forcar: bool = False):
        if not forcar:
            if etapa >= 2 and self.df is None:
                messagebox.showwarning(
                    self._t("dialog.title.step_locked", "Etapa bloqueada"),
                    self._t("dialog.step2_blocked", "Para avançar para Configuração, selecione e carregue um arquivo."),
                )
                self.status_resumo_var.set(self._t("status.step2_blocked", "Etapa 2 bloqueada: selecione um arquivo primeiro."))
                return
            if etapa >= 3 and (self.df is None or not self.column_combo.get()):
                messagebox.showwarning(
                    self._t("dialog.title.step_locked", "Etapa bloqueada"),
                    self._t("dialog.step3_blocked", "Para avançar para Ação, selecione uma coluna válida."),
                )
                self.status_resumo_var.set(self._t("status.step3_blocked", "Etapa 3 bloqueada: selecione uma coluna válida."))
                return

        self.etapa_atual = etapa
        for frame in (
            self.sections.dados_frame,
            self.sections.config_frame,
            self.sections.acao_status_frame,
        ):
            if frame is not None:
                frame.pack_forget()

        if etapa == 1:
            self.dados_frame.pack(fill="x", pady=(0, self.space_sm))
            self.select_button.focus_set()
        elif etapa == 2:
            self.config_frame.pack(fill="x", pady=(0, self.space_sm))
            self.formato_combo.focus_set()
        else:
            self.acao_status_frame.pack(fill="x")
            self.generate_button.focus_set()

        self._atualizar_stepper_visual()

    def _ao_selecionar_coluna(self, _e=None):
        self.solicitar_atualizacao_preview()
        self._atualizar_stepper_visual()

    def _ao_alterar_formato_saida(self, _e=None):
        formatos_disponiveis = set(self.formato_combo.cget("values"))
        if self.formato_saida.get() not in formatos_disponiveis:
            self.formato_saida.set("png")
            self.formato_combo.set("png")
        self._atualizar_controles_impressao()
        self._aplicar_estado_ui()

    def _ao_alterar_tipo_codigo(self):
        self._atualizar_controles_tipo_codigo()
        self.solicitar_atualizacao_preview()

    def _ao_alterar_modelo_barcode(self, _e=None):
        chave = self.barcode_label_to_key.get(self.barcode_model_combo.get())
        if chave:
            self.barcode_model.set(chave)
        self.solicitar_atualizacao_preview()

    def _atualizar_controles_tipo_codigo(self):
        if not self.barcode_disponivel:
            self.tipo_codigo.set("qrcode")
            self.barcode_model_combo.configure(state="disabled")
            return
        if self.tipo_codigo.get() == "barcode":
            self.barcode_model_combo.configure(state="readonly")
            if self.barcode_model.get() in self.barcode_key_to_label:
                self.barcode_model_combo.set(self.barcode_key_to_label[self.barcode_model.get()])
        else:
            self.barcode_model_combo.configure(state="disabled")

    def _aplicar_disponibilidade_dependencias(self):
        if hasattr(self, "tipo_barcode_radio"):
            self.tipo_barcode_radio.configure(state="normal" if self.barcode_disponivel else "disabled")
        if not self.barcode_disponivel:
            self.tipo_codigo.set("qrcode")

        formatos_disponiveis = ["png", "zip", "svg"]
        if self.pdf_export_disponivel:
            formatos_disponiveis = ["pdf", "png", "zip", "svg", "imprimir"]
        self.formato_combo.configure(values=formatos_disponiveis)
        if self.formato_saida.get() not in formatos_disponiveis:
            self.formato_saida.set("png")
            self.formato_combo.set("png")

    def _listar_impressoras_windows(self):
        if not sys.platform.startswith("win"):
            return [], ""
        comando = (
            "Get-CimInstance Win32_Printer | "
            "Select-Object Name,Default | "
            "ConvertTo-Json -Compress"
        )
        saida = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", comando],
            stderr=subprocess.STDOUT,
            timeout=10,
            text=True,
        ).strip()
        if not saida:
            return [], ""

        dados = json.loads(saida)
        if isinstance(dados, dict):
            dados = [dados]

        impressoras = []
        padrao = ""
        for item in dados:
            nome = str(item.get("Name", "")).strip()
            if not nome:
                continue
            impressoras.append(nome)
            if bool(item.get("Default")):
                padrao = nome

        impressoras = sorted(set(impressoras), key=lambda nome: nome.lower())
        return impressoras, padrao

    def atualizar_lista_impressoras(self, notificar=False):
        if not sys.platform.startswith("win"):
            self.impressora_combo.configure(state="disabled", values=[])
            self.impressora_status_var.set(self._t("print.windows_only", "Impressão integrada disponível apenas no Windows."))
            return
        try:
            impressoras, padrao = self._listar_impressoras_windows()
        except Exception as exc:
            self.impressora_combo.configure(state="disabled", values=[])
            self.impressora_status_var.set(self._t("print.query_error", "Não foi possível consultar as impressoras instaladas."))
            if notificar:
                messagebox.showerror(self._t("dialog.title.print", "Impressão"), self._t("print.list_error", "Erro ao listar impressoras:\n{erro}", erro=exc))
            return

        self.impressora_combo.configure(values=impressoras)
        if impressoras:
            atual = self.impressora_var.get().strip()
            selecionada = atual if atual in impressoras else (padrao or impressoras[0])
            self.impressora_var.set(selecionada)
            self.impressora_combo.configure(state="readonly")
            self.impressora_status_var.set(
                self._t(
                    "print.found_summary",
                    "{quantidade} impressora(s) encontrada(s). Padrão: {padrao}.",
                    quantidade=len(impressoras),
                    padrao=padrao or self._t("print.not_defined", "não definida"),
                )
            )
            if notificar:
                messagebox.showinfo(self._t("dialog.title.print", "Impressão"), self._t("print.list_updated", "Lista de impressoras atualizada."))
        else:
            self.impressora_var.set("")
            self.impressora_combo.configure(state="disabled")
            self.impressora_status_var.set(self._t("print.none_available", "Nenhuma impressora disponível no sistema."))
            if notificar:
                messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.none_found", "Nenhuma impressora foi encontrada."))

    def _atualizar_controles_impressao(self):
        formato_impressao = self.formato_saida.get() == "imprimir"
        bloqueado = self.estado_atual in {EstadoAplicacao.LOADING, EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}
        if formato_impressao:
            self.impressao_frame.grid()
        else:
            self.impressao_frame.grid_remove()

        if not formato_impressao:
            return

        self.copias_spin.configure(state="disabled" if bloqueado else "normal")
        if not bloqueado and bool(self.impressora_combo.cget("values")):
            self.impressora_combo.configure(state="readonly")
        elif bloqueado:
            self.impressora_combo.configure(state="disabled")

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
            caminho_abrivel = bool(caminho and not str(caminho).startswith("impressora:"))
            self.abrir_pasta_button.configure(state="normal" if caminho_abrivel else "disabled")
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
            messagebox.showerror(
                self._t("message.error", "Erro"),
                self._t("error.open_output_folder", "Não foi possível abrir a pasta de saída:\n{erro}", erro=exc),
            )

    def atualizar_controles_formato(self):
        if self.modo.get() == "texto":
            self.numerico_controls.grid_forget()
            self.texto_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)
        else:
            self.texto_controls.grid_forget()
            self.numerico_controls.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=3)

    def _transicionar_estado(self, novo_estado: EstadoAplicacao):
        self.estado_atual = novo_estado
        self._aplicar_estado_ui()

    def _aplicar_estado_ui(self):
        bloqueado = self.estado_atual in {EstadoAplicacao.LOADING, EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}
        pode_cancelar = self.estado_atual in {EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}
        pode_gerar = (not bloqueado) and self.df is not None and bool(self.column_combo.get())
        pode_teste_impressao = pode_gerar and self.formato_saida.get() == "imprimir"

        self.select_button.configure(state="disabled" if bloqueado else "normal")
        self.generate_button.configure(state="normal" if pode_gerar else "disabled")
        self.test_print_button.configure(state="normal" if pode_teste_impressao else "disabled")
        self.cancel_button.configure(
            state="normal" if pode_cancelar else "disabled",
            style="Danger.TButton" if pode_cancelar else "Secondary.TButton",
        )

        if self.df is not None and self._obter_colunas(self.df):
            self.column_combo.configure(state="readonly")
        else:
            self.column_combo.configure(state="disabled")

        self._atualizar_controles_impressao()
        if hasattr(self, "barcode_model_combo"):
            barcode_ativo = self.barcode_disponivel and self.tipo_codigo.get() == "barcode"
            estado_barcode = "readonly" if (barcode_ativo and not bloqueado) else "disabled"
            self.barcode_model_combo.configure(state=estado_barcode)
        self._atualizar_stepper_visual()

    def cancelar_operacao(self):
        if self.estado_atual not in {EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}:
            return
        self.cancelar_evento.set()
        self._transicionar_estado(EstadoAplicacao.CANCELLING)
        self.progress_label_var.set(self._t("progress.cancelling", "Cancelando operação..."))

    def selecionar_arquivo(self):
        if self.estado_atual in {EstadoAplicacao.LOADING, EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}:
            return

        caminho = filedialog.askopenfilename(
            title=self._t("filedialog.open_data", "Selecione CSV ou Excel"),
            filetypes=[(self._t("filedialog.data_files", "Arquivos de dados"), "*.csv *.xlsx")],
        )
        if not caminho:
            return

        self._transicionar_estado(EstadoAplicacao.LOADING)
        self.progress_bar.stop()
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar["value"] = 0
        self.progress_bar.start(10)
        self.progress_label_var.set(self._t("progress.loading_file", "Carregando arquivo, aguarde..."))
        self.status_resumo_var.set(self._t("status.loading_spreadsheet", "Carregando dados da planilha..."))
        self.progress_frame.pack(fill="x")

        self.logger.info("Iniciando carregamento de arquivo", extra={"event": "load_start", "operation": "load", "path": caminho})
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
        return self.controller.formatar_excecao(exc, contexto)

    def _carregar_tabela(self, caminho):
        """Carrega CSV/XLSX com fallback quando pandas/numpy não estiverem disponíveis."""
        return self.controller.carregar_tabela(caminho)

    def _obter_colunas(self, tabela):
        return self.controller.obter_colunas(tabela)

    def _obter_valores_coluna(self, tabela, coluna):
        return self.controller.obter_valores_coluna(tabela, coluna)

    def _parse_float_input(self, valor, nome_campo: str) -> float:
        if isinstance(valor, str):
            texto = valor.strip().replace(",", ".")
        else:
            texto = str(valor)
        try:
            return float(texto)
        except (TypeError, ValueError) as exc:
            raise ValueError(self._t("validation.invalid_number", "Valor inválido para {campo}: {valor}", campo=nome_campo, valor=valor)) from exc

    def _build_config(self) -> GeracaoConfig:
        if self.tipo_codigo.get() == "barcode" and not self.barcode_disponivel:
            raise ValueError("Código de barras indisponível neste ambiente (dependência renderPM ausente).")
        if self.formato_saida.get() in {"pdf", "imprimir"} and not self.pdf_export_disponivel:
            raise ValueError("Formato PDF/Impressão indisponível neste ambiente (dependência reportlab ausente).")
        return GeracaoConfig(
            qr_width_cm=self._parse_float_input(self.qr_width_cm.get(), self._t("labels.qr_width", "Largura QR (cm)")),
            qr_height_cm=self._parse_float_input(self.qr_height_cm.get(), self._t("labels.qr_height", "Altura QR (cm)")),
            barcode_width_cm=self._parse_float_input(self.barcode_width_cm.get(), self._t("labels.barcode_width", "Largura Código de Barras (cm)")),
            barcode_height_cm=self._parse_float_input(self.barcode_height_cm.get(), self._t("labels.barcode_height", "Altura Código de Barras (cm)")),
            keep_qr_ratio=bool(self.keep_qr_ratio.get()),
            keep_barcode_ratio=bool(self.keep_barcode_ratio.get()),
            foreground=self.qr_foreground_color.get(),
            background=self.qr_background_color.get(),
            tipo_codigo=self.tipo_codigo.get(),
            barcode_model=self.barcode_model.get(),
            modo=self.modo.get(),
            prefixo=self.prefixo_numerico.get(),
            sufixo=self.sufixo_numerico.get(),
            max_codigos_por_lote=self.max_codigos_por_lote,
            max_tamanho_dado=self.max_tamanho_dado,
        )

    def _validar_parametros_geracao(self, codigos, cfg: GeracaoConfig | None = None):
        cfg = cfg or self._build_config()
        return self.controller.validar_parametros_geracao(codigos, cfg)

    def _sanitizar_nome_arquivo(self, nome: str, fallback: str) -> str:
        return self.controller.sanitizar_nome_arquivo(nome, fallback)

    def _normalizar_dado(self, valor: str, cfg: GeracaoConfig | None = None) -> str:
        cfg = cfg or self._build_config()
        return self.controller.normalizar_dado(valor, cfg)

    def _gerar_imagem_obj(self, dado: str, cfg: GeracaoConfig | None = None) -> Image.Image:
        cfg = cfg or self._build_config()
        return self.controller.gerar_imagem_obj(dado, cfg)

    def _extrair_codigos_preview(self):
        if self.df is None or not self.column_combo.get():
            return []
        try:
            cfg = self._build_config()
            return self.controller.extrair_codigos_preview(
                self.df,
                self.column_combo.get(),
                cfg,
                self.max_itens_preview_pagina,
            )
        except ValueError:
            return []

    def _gerar_preview_documento(self, codigos, cfg: GeracaoConfig) -> Image.Image:
        preset = self.preview_preset.get()
        if preset == "Etiqueta 8x10.5 cm":
            largura, altura = int(80 * mm), int(105 * mm)
        elif preset == "Etiqueta 60x40 mm":
            largura, altura = int(60 * mm), int(40 * mm)
        else:
            largura, altura = map(int, A4)

        margem_cm = max(0.2, self._parse_float_input(self.preview_margin_cm.get(), self._t("labels.preview_margin", "Margem (cm)")))
        espaco_cm = max(0.1, self._parse_float_input(self.preview_spacing_cm.get(), self._t("labels.preview_spacing", "Espaçamento (cm)")))
        margem_px = int(margem_cm * 10 * mm)
        espaco_px = int(espaco_cm * 10 * mm)

        # Evita geometria inválida (x1 < x0 / y1 < y0) em etiquetas pequenas
        # ou quando margem é maior que metade da área útil.
        margem_limite = max(1, min((largura - 2) // 2, (altura - 2) // 2))
        margem_px = min(margem_px, margem_limite)

        fundo = Image.new("RGB", (largura + 120, altura + 120), "#e5e7eb")
        draw = ImageDraw.Draw(fundo)
        draw.rounded_rectangle((70, 70, largura + 90, altura + 90), radius=10, fill="#cbd5e1")

        preview = Image.new("RGB", (largura, altura), "white")
        draw_preview = ImageDraw.Draw(preview)

        x = margem_px
        y = margem_px
        x2 = largura - x
        y2 = altura - y
        # Coordenadas defensivas para evitar ValueError do Pillow em presets pequenos.
        x0, x1 = sorted((x, x2))
        y0, y1 = sorted((y, y2))
        x0 = max(0, min(x0, largura - 1))
        x1 = max(0, min(x1, largura - 1))
        y0 = max(0, min(y0, altura - 1))
        y1 = max(0, min(y1, altura - 1))
        item_largura = max(24, int(cfg.qr_width_cm * 10 * mm))
        item_altura = max(24, int(cfg.qr_height_cm * 10 * mm))
        if cfg.tipo_codigo == "barcode":
            item_largura = max(24, int(cfg.barcode_width_cm * 10 * mm))
            item_altura = max(24, int(cfg.barcode_height_cm * 10 * mm))

        if x1 >= x0 and y1 >= y0:
            draw_preview.rectangle((x0, y0, x1, y1), outline="#9ca3af", width=2)

        pixels_por_cm = mm * 10
        largura_util = max(1, x1 - x0)
        for cm in range(0, int(largura_util / pixels_por_cm) + 1, 5):
            px = int(x0 + cm * pixels_por_cm)
            if px > x1:
                break
            if y0 >= 8:
                draw_preview.line((px, y0 - 8, px, y0), fill="#6b7280", width=1)
            draw_preview.text((px + 2, max(0, y0 - 24)), f"{cm}cm", fill="#6b7280")

        x_cursor = x0
        y_cursor = y0

        for codigo in codigos:
            img = self._gerar_imagem_obj(self._normalizar_dado(codigo, cfg), cfg)
            img = img.resize((item_largura, item_altura))
            preview.paste(img, (x_cursor, y_cursor))

            x_cursor += item_largura + espaco_px
            if x_cursor + item_largura > x1:
                x_cursor = x0
                y_cursor += item_altura + espaco_px
            if y_cursor + item_altura > y1:
                break

        fundo.paste(preview, (60, 60))
        return fundo

    def atualizar_preview(self):
        try:
            cfg = self._build_config()
            codigos_preview = self._extrair_codigos_preview()
            if codigos_preview:
                img = self._gerar_preview_documento(codigos_preview, cfg)
            else:
                _ = self.controller.gerar_amostra_preview(cfg)
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
                messagebox.showwarning(self._t("dialog.title.missing_dependency", "Dependência opcional ausente"), str(exc))
                self._preview_backend_error_shown = True
        except Exception as exc:
            self.preview_label.configure(image="", text="Falha ao gerar pré-visualização")
            self.logger.exception(
                "Erro inesperado no preview",
                extra={"event": "preview_error", "operation": "preview", "erro": str(exc)},
            )
            if not self._preview_backend_error_shown:
                messagebox.showwarning(self._t("dialog.title.preview", "Pré-visualização"), self._t("preview.update_error", "Não foi possível atualizar o preview:\n{erro}", erro=exc))
                self._preview_backend_error_shown = True

    def solicitar_atualizacao_preview(self, *_args):
        if self._preview_after_id is not None:
            self.root.after_cancel(self._preview_after_id)
        self._preview_after_id = self.root.after(self.preview_debounce_ms, self._executar_preview_debounced)

    def _executar_preview_debounced(self):
        self._preview_after_id = None
        self.atualizar_preview()

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

    def gerar_pdf(self, codigos, caminho_pdf, emitir_sucesso=True):
        try:
            cfg = self._build_config()
            pdf_canvas, image_reader_cls = _obter_modulos_pdf_reportlab()
            pdf = pdf_canvas.Canvas(caminho_pdf, pagesize=A4)
            largura_pagina, altura_pagina = A4

            x = 20 * mm
            y = altura_pagina - 20 * mm
            largura_item = cfg.qr_width_cm * 10 * mm
            altura_item = cfg.qr_height_cm * 10 * mm
            if cfg.tipo_codigo == "barcode":
                largura_item = cfg.barcode_width_cm * 10 * mm
                altura_item = cfg.barcode_height_cm * 10 * mm
            margem = 10 * mm
            y -= altura_item
            total = len(codigos)

            for i, codigo in enumerate(codigos, start=1):
                if self.cancelar_evento.is_set():
                    raise OperacaoCancelada("Operação cancelada pelo usuário.")
                imagem = self._gerar_imagem_obj(self._normalizar_dado(codigo, cfg), cfg)
                buffer = io.BytesIO()
                imagem.save(buffer, format="PNG")
                buffer.seek(0)
                image_reader = image_reader_cls(buffer)

                pdf.drawImage(image_reader, x, y, width=largura_item, height=altura_item, preserveAspectRatio=True)
                self.fila.put({"tipo": "progresso", "atual": i, "total": total, "codigo": codigo})

                x += largura_item + margem
                if x + largura_item > largura_pagina - 20 * mm:
                    x = 20 * mm
                    y -= altura_item + margem

                if y < 20 * mm:
                    pdf.showPage()
                    x = 20 * mm
                    y = altura_pagina - 20 * mm - altura_item

            pdf.save()
            if emitir_sucesso:
                self.logger.info("PDF gerado com sucesso", extra={"event": "generate_done", "operation": "pdf", "path": caminho_pdf, "total": total})
                self.fila.put({"tipo": "sucesso", "caminho": caminho_pdf})
        except (OSError, ValueError, RuntimeError) as exc:
            raise RuntimeError(self._formatar_excecao(exc, "Erro ao gerar PDF")) from exc

    def imprimir_codigos(self, codigos):
        if not sys.platform.startswith("win"):
            raise RuntimeError("A impressão integrada está disponível apenas no Windows.")

        try:
            copias = max(1, int(self.copias_impressao.get()))
        except Exception as exc:
            raise RuntimeError("Quantidade de cópias inválida.") from exc

        impressora = self.impressora_var.get().strip()
        pasta_tmp = tempfile.mkdtemp(prefix="qr_print_")
        self._arquivos_temporarios_impressao.append(pasta_tmp)
        self.gerar_imagens(codigos, "png", pasta_tmp, emitir_sucesso=False)

        arquivos_png = [
            os.path.join(pasta_tmp, nome)
            for nome in sorted(os.listdir(pasta_tmp))
            if nome.lower().endswith(".png")
        ]
        if not arquivos_png:
            raise RuntimeError("Nenhum arquivo foi gerado para impressão.")

        for _ in range(copias):
            for caminho_imagem in arquivos_png:
                if self.cancelar_evento.is_set():
                    raise OperacaoCancelada("Operação cancelada pelo usuário.")
                self._imprimir_png_windows(caminho_imagem, impressora)
                time.sleep(0.2)

        destino = f"impressora:{impressora or 'padrão do sistema'}"
        self.logger.info(
            "Impressão enviada com sucesso",
            extra={"event": "print_done", "operation": "print", "path": destino, "total": len(codigos), "copias": copias},
        )
        self.fila.put(
            {
                "tipo": "sucesso",
                "caminho": destino,
                "descricao": f"Envio para impressão concluído em {impressora or 'impressora padrão'} ({copias} cópia(s)).",
            }
        )

    def _imprimir_png_windows(self, caminho_imagem: str, impressora: str):
        if impressora:
            cmd = ["mspaint.exe", "/pt", caminho_imagem, impressora]
        else:
            cmd = ["mspaint.exe", "/p", caminho_imagem]
        try:
            subprocess.run(cmd, check=True, timeout=30)
        except FileNotFoundError as exc:
            raise RuntimeError("Não foi possível localizar o mspaint.exe para realizar a impressão.") from exc
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"Falha ao enviar imagem para impressão (código {exc.returncode}).") from exc
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("Tempo excedido ao enviar imagem para a impressora.") from exc

    def _iniciar_progresso(self, total, invalidos=0, destino="", formato=""):
        self.cancelar_evento.clear()
        self._transicionar_estado(EstadoAplicacao.GENERATING)
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate", maximum=max(total, 1))
        self.progress_bar["value"] = 0
        self.progress_label_var.set(self._t("progress.generating_start", "Gerando 0/{total}...", total=total))
        self.status_resumo_var.set(self._t("status.processing_records", "Processando {total} registro(s)...", total=total))
        self._inicio_geracao_ts = time.perf_counter()
        self._total_planejado = total
        self._processados_atuais = 0
        self._invalidos_ultima_geracao = invalidos
        self._atualizar_resumo_painel(processado=0, ignorados=invalidos, duracao=0, caminho=destino, job_id="")
        self._formato_execucao_atual = formato or self.formato_saida.get()
        self._registrar_job(formato=formato or self.formato_saida.get(), destino=str(destino), total_validos=total, invalidos=invalidos)
        self.progress_frame.pack(fill="x")

    def _finalizar_progresso(self, estado_final: EstadoAplicacao | None = None):
        duracao = None
        if self._inicio_geracao_ts is not None:
            duracao = time.perf_counter() - self._inicio_geracao_ts

        self.cancelar_evento.clear()
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        self._atualizar_resumo_painel(
            processado=self._processados_atuais,
            ignorados=self._invalidos_ultima_geracao,
            duracao=duracao,
        )
        self._inicio_geracao_ts = None
        if estado_final is None:
            estado_final = EstadoAplicacao.READY if self.df is not None else EstadoAplicacao.IDLE
        self._transicionar_estado(estado_final)

    def verificar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg["tipo"] == "progresso":
                    atual = msg.get("atual", 0)
                    total = msg.get("total", 1)
                    self.progress_bar.configure(maximum=max(total, 1), style="App.Horizontal.TProgressbar")
                    self.progress_bar["value"] = atual
                    self.progress_label_var.set(
                        self._t("progress.generating_item", "Gerando {atual}/{total}: {codigo}", atual=atual, total=total, codigo=msg.get("codigo", ""))
                    )
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
                    self._finalizar_progresso(EstadoAplicacao.READY)
                    self.status_resumo_var.set(self._t("status.completed_output", "Concluído com sucesso. Saída: {caminho}", caminho=msg.get("caminho", "")))
                    messagebox.showinfo(
                        self._t("message.success", "Sucesso"),
                        msg.get(
                            "descricao",
                            self._t("info.generated_files_in", "Arquivo(s) gerado(s) em: {caminho}", caminho=msg.get("caminho", "")),
                        ),
                    )
                elif msg["tipo"] == "erro":
                    self._atualizar_resumo_painel(caminho=self._ultimo_destino_saida)
                    self.progress_bar.configure(style="Error.Horizontal.TProgressbar")
                    self._registrar_metricas_execucao("error", erro=msg.get("msg", ""))
                    self._finalizar_job("error", erro=msg.get("msg", ""))
                    self._finalizar_progresso(EstadoAplicacao.ERROR)
                    self.status_resumo_var.set(self._t("status.generation_failed", "Falha durante a geração. Verifique os detalhes do erro."))
                    detalhe = msg.get("detalhe", "")
                    erro_msg = msg.get("msg", self._t("error.generation_failed", "Falha durante a geração."))
                    if detalhe:
                        erro_msg = self._t("error.technical_details", "{erro}\n\nDetalhes técnicos:\n{detalhe}", erro=erro_msg, detalhe=detalhe)
                    messagebox.showerror(self._t("message.error", "Erro"), erro_msg)
                elif msg["tipo"] == "cancelado":
                    self._atualizar_resumo_painel(caminho=self._ultimo_destino_saida)
                    self.progress_bar.configure(style="App.Horizontal.TProgressbar")
                    self._registrar_metricas_execucao("cancelled", erro=msg.get("msg", ""))
                    self._finalizar_job("cancelled", erro=msg.get("msg", ""))
                    self._finalizar_progresso(EstadoAplicacao.READY if self.df is not None else EstadoAplicacao.IDLE)
                    self.status_resumo_var.set(self._t("status.operation_cancelled", "Operação cancelada pelo usuário."))
                    messagebox.showinfo(
                        self._t("dialog.title.cancelled", "Cancelado"),
                        msg.get("msg", self._t("info.operation_cancelled", "Operação cancelada.")),
                    )
                elif msg["tipo"] == "carregamento_sucesso":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self.df = msg["tabela"]
                    self.arquivo_fonte = msg["caminho"]
                    colunas = self._obter_colunas(self.df)
                    self.column_combo.configure(values=colunas)
                    self.status_resumo_var.set(
                        f"Arquivo carregado: {os.path.basename(self.arquivo_fonte)} ({len(colunas)} coluna(s))."
                    )
                    self._atualizar_resumo_painel(processado=0, ignorados=0, caminho="", job_id="")
                    if colunas:
                        self.column_combo.set(colunas[0])
                    self.solicitar_atualizacao_preview()
                    self._transicionar_estado(EstadoAplicacao.READY)
                elif msg["tipo"] == "carregamento_erro":
                    self.progress_bar.stop()
                    self.progress_frame.pack_forget()
                    self.column_combo.configure(state="disabled", values=[])
                    self._transicionar_estado(EstadoAplicacao.ERROR)
                    self.status_resumo_var.set(self._t("status.file_load_failed", "Falha ao carregar arquivo. Tente novamente."))
                    self._atualizar_resumo_painel(processado=0, ignorados=0, caminho="", job_id="")
                    messagebox.showerror(
                        self._t("message.error", "Erro"),
                        self._t("error.open_file_failed", "Não foi possível abrir o arquivo: {erro}", erro=msg.get("msg", "")),
                    )
        except queue.Empty:
            pass

        self.root.after(100, self.verificar_fila)

    def _executar_geracao(self, codigos, formato, destino):
        try:
            if formato == "pdf":
                self.gerar_pdf(codigos, destino)
            elif formato == "zip":
                self.gerar_zip(codigos, destino)
            elif formato == "imprimir":
                self.imprimir_codigos(codigos)
            else:
                self.gerar_imagens(codigos, formato, destino)
        except OperacaoCancelada as exc:
            self.logger.info("Geração cancelada", extra={"event": "generate_cancel", "operation": formato, "path": str(destino)})
            self.fila.put({"tipo": "cancelado", "msg": str(exc)})
        except Exception as exc:
            self.logger.exception("Falha na geração", extra={"event": "generate_error", "operation": formato, "path": str(destino), "erro": str(exc)})
            self.fila.put({"tipo": "erro", "msg": str(exc), "detalhe": traceback.format_exc(limit=3)})

    def gerar_a_partir_da_tabela(self):
        if self.estado_atual in {EstadoAplicacao.LOADING, EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}:
            return
        if self.df is None or not self.column_combo.get():
            messagebox.showwarning(self._t("message.warning", "Aviso"), self._t("warning.select_file_and_column", "Selecione um arquivo e uma coluna."))
            return

        try:
            cfg = self._build_config()
            codigos, invalidos = self.controller.preparar_codigos(self.df, self.column_combo.get(), cfg)
        except ValueError as exc:
            messagebox.showwarning(self._t("dialog.title.validation", "Validação"), str(exc))
            return

        if invalidos:
            messagebox.showwarning(
                self._t("dialog.title.validation", "Validação"),
                self._t(
                    "warning.invalid_records_ignored",
                    "{invalidos} registro(s) foram ignorados por não atenderem aos limites de entrada.",
                    invalidos=invalidos,
                ),
            )

        formato = self.formato_saida.get()
        destino = None
        if formato == "pdf":
            destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        elif formato == "zip":
            destino = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
        elif formato == "imprimir":
            if not sys.platform.startswith("win"):
                messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.windows_only", "A impressão integrada está disponível apenas no Windows."))
                return
            try:
                copias = int(self.copias_impressao.get())
            except Exception:
                messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.invalid_copies", "Informe uma quantidade de cópias válida."))
                return
            if copias < 1:
                messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.copies_gt_zero", "A quantidade de cópias deve ser maior que zero."))
                return
            impressora = self.impressora_var.get().strip()
            destino = f"impressora:{impressora or 'padrão do sistema'}"
        else:
            destino = filedialog.askdirectory()

        if not destino:
            return

        self.logger.info("Iniciando geração", extra={"event": "generate_start", "operation": formato, "path": str(destino), "total": len(codigos)})
        self._iniciar_progresso(len(codigos), invalidos=invalidos, destino=str(destino), formato=formato)
        worker = threading.Thread(target=self._executar_geracao, args=(codigos, formato, destino), daemon=True)
        worker.start()

    def imprimir_teste(self):
        if self.estado_atual in {EstadoAplicacao.LOADING, EstadoAplicacao.GENERATING, EstadoAplicacao.CANCELLING}:
            return
        if self.df is None or not self.column_combo.get():
            messagebox.showwarning(self._t("message.warning", "Aviso"), self._t("warning.select_file_and_column", "Selecione um arquivo e uma coluna."))
            return
        if self.formato_saida.get() != "imprimir":
            messagebox.showwarning(
                self._t("dialog.title.print", "Impressão"),
                self._t("print.select_output_print", "Selecione o formato de saída 'imprimir' para usar o teste."),
            )
            return
        if not sys.platform.startswith("win"):
            messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.windows_only", "A impressão integrada está disponível apenas no Windows."))
            return

        try:
            copias = int(self.copias_impressao.get())
        except Exception:
            messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.invalid_copies", "Informe uma quantidade de cópias válida."))
            return
        if copias < 1:
            messagebox.showwarning(self._t("dialog.title.print", "Impressão"), self._t("print.copies_gt_zero", "A quantidade de cópias deve ser maior que zero."))
            return

        try:
            cfg = self._build_config()
            codigos, invalidos = self.controller.preparar_codigos(self.df, self.column_combo.get(), cfg)
        except ValueError as exc:
            messagebox.showwarning(self._t("dialog.title.validation", "Validação"), str(exc))
            return

        if not codigos:
            messagebox.showwarning(
                self._t("dialog.title.validation", "Validação"),
                self._t("validation.no_valid_codes_for_print", "Nenhum código válido encontrado para impressão."),
            )
            return

        codigo_teste = [codigos[0]]
        impressora = self.impressora_var.get().strip()
        destino = f"impressora:{impressora or 'padrão do sistema'} (teste)"

        self.logger.info(
            "Iniciando impressão de teste",
            extra={"event": "print_test_start", "operation": "imprimir_teste", "path": destino, "codigo": codigo_teste[0]},
        )
        self._iniciar_progresso(len(codigo_teste), invalidos=invalidos, destino=destino, formato="imprimir")
        worker = threading.Thread(target=self._executar_geracao, args=(codigo_teste, "imprimir", destino), daemon=True)
        worker.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()
