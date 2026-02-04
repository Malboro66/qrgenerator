import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser
import pandas as pd
import qrcode
import qrcode.image.svg
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import io
import os
import threading
import queue
import re
import time  # Importante para Throttling
from PIL import Image, ImageDraw, ImageTk
import zipfile


class QRGeneratorService:
    """
    Classe de Serviço (Lógica de Negócio).
    Responsável pela geração de QR Codes, manipulação de arquivos,
    sanitização de nomes e exportação (PDF, PNG, SVG, ZIP).
    Não possui dependências diretas da UI (Tkinter).
    """
    def __init__(self):
        # Configurações Padrão
        self.qr_size = 200
        self.qr_foreground_color = "black"
        self.qr_background_color = "white"
        self.logo_path = ""
        
        # Constantes PDF
        self.COLUNAS_PDF = 3
        self.LINHAS_PDF = 4
        self.MARGEM_CM = 1.0
        
        # --- MELHORIA: Cache do QR Code Base ---
        # Armazena o objeto QRCode (matriz) calculado para evitar recálculo pesado
        self._cached_qr_object = None
        self._cache_data = None

    def atualizar_config(self, size, fg_color, bg_color, logo_path):
        """Atualiza as configurações de geração."""
        self.qr_size = int(size)
        self.qr_foreground_color = fg_color
        self.qr_background_color = bg_color
        self.logo_path = logo_path

    def _sanitizar_nome(self, nome):
        """Remove caracteres inválidos de nomes de arquivos."""
        if not nome:
            return "sem_nome"
        return re.sub(r'[\\/*?:"<>|]', "_", str(nome)).strip()

    def _criar_objeto_qr_com_cache(self, data):
        """
        Cria o objeto QRCode, utilizando cache se os dados forem os mesmos.
        Isso economiza processamento pesado ao mudar apenas cores ou logo.
        """
        if data == self._cache_data and self._cached_qr_object is not None:
            return self._cached_qr_object
        
        # Se os dados mudaram, precisamos recalcular a matriz
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)
        
        # Atualiza o cache
        self._cached_qr_object = qr
        self._cache_data = data
        return qr

    def gerar_imagem_qr(self, data):
        """Gera um objeto de imagem PIL (Raster) pronto para uso ou preview."""
        # Usa o método com cache para performance
        qr = self._criar_objeto_qr_com_cache(data)
        
        img = qr.make_image(
            fill_color=self.qr_foreground_color,
            back_color=self.qr_background_color,
        )
        img = img.resize((self.qr_size, self.qr_size), Image.Resampling.LANCZOS)
        
        if self.logo_path:
            try:
                logo = Image.open(self.logo_path)
                logo_size = int(self.qr_size * 0.2)
                logo = logo.resize((logo_size, logo_size), Image.Resampling.LANCZOS)
                pos = ((img.size[0] - logo.size[0]) // 2, (img.size[1] - logo.size[1]) // 2)
                img.paste(logo, pos, logo if logo.mode == "RGBA" else None)
            except Exception:
                pass
        return img

    def _desenhar_qr_no_pdf(self, pdf, img, x, y, codigo, largura_cm):
        buffer_img = io.BytesIO()
        img.save(buffer_img, format="PNG")
        buffer_img.seek(0)
        width_pt = largura_cm * cm
        pdf.drawImage(ImageReader(buffer_img), x, y, width=width_pt, height=width_pt, mask='auto')
        pdf.setFont("Helvetica", 8)
        texto_largura = pdf.stringWidth(codigo, "Helvetica", 8)
        pdf.drawString(x + (width_pt - texto_largura) / 2, y - 12, codigo)

    def _deveria_atualizar_progresso(self, indice, total, last_update_time, throttle_segundos=0.2):
        """
        Auxiliar para Throttling.
        Retorna True se deve atualizar a UI baseado no tempo decorrido ou se é o último item.
        """
        agora = time.time()
        # Atualiza se passou o tempo limite OU se for o último item (para garantir 100%)
        if (agora - last_update_time) > throttle_segundos or indice == total - 1:
            return True, agora
        return False, last_update_time

    def gerar_pdf(self, codigos, caminho_pdf, fila_progresso):
        pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
        largura_pagina, altura_pagina = A4
        margem = self.MARGEM_CM * cm
        area_util_largura = largura_pagina - (2 * margem)
        largura_coluna = area_util_largura / self.COLUNAS_PDF
        area_util_altura = altura_pagina - (2 * margem)
        altura_linha = area_util_altura / self.LINHAS_PDF
        tamanho_celula = min(largura_coluna, altura_linha)
        qr_por_pagina = self.COLUNAS_PDF * self.LINHAS_PDF

        try:
            last_update = time.time()
            for indice, codigo in enumerate(codigos):
                img = self.gerar_imagem_qr(codigo)
                
                idx_grid = indice % qr_por_pagina
                linha = idx_grid // self.COLUNAS_PDF
                coluna = idx_grid % self.COLUNAS_PDF
                
                x = margem + (coluna * tamanho_celula)
                y = margem + ((self.LINHAS_PDF - linha - 1) * tamanho_celula)
                y_desenho = y + 0.5 * cm 
                
                self._desenhar_qr_no_pdf(pdf, img, x, y_desenho, codigo, (tamanho_celula - 0.5 * cm)/cm)

                if (indice + 1) % qr_por_pagina == 0 and indice < len(codigos) - 1:
                    pdf.showPage()
                
                # --- MELHORIA: Throttling de Progresso ---
                should_update, last_update = self._deveria_atualizar_progresso(indice, len(codigos), last_update)
                if should_update:
                    fila_progresso.put({"tipo": "progresso", "atual": indice, "total": len(codigos)})
            
            pdf.save()
            fila_progresso.put({"tipo": "sucesso", "caminho": caminho_pdf})
        except Exception as e:
            fila_progresso.put({"tipo": "erro", "mensagem": f"Erro ao gerar PDF: {str(e)}"})

    def gerar_imagens(self, codigos, formato, diretorio, fila_progresso):
        try:
            total = len(codigos)
            last_update = time.time()
            
            for i, codigo in enumerate(codigos):
                nome_arquivo = self._sanitizar_nome(codigo)
                caminho_arquivo = os.path.join(diretorio, f"{nome_arquivo}.{formato}")
                
                if formato == "png":
                    img = self.gerar_imagem_qr(codigo)
                    img.save(caminho_arquivo, format="PNG")
                elif formato == "svg":
                    # SVG não usa cache de objeto PIL da mesma forma, mas o cálculo do QR é rápido
                    qr = self._criar_objeto_qr_com_cache(codigo)
                    img = qr.make_image(image_factory=qrcode.image.svg.SvgImage)
                    img.save(caminho_arquivo)
                
                # --- MELHORIA: Throttling de Progresso ---
                should_update, last_update = self._deveria_atualizar_progresso(i, total, last_update)
                if should_update:
                    fila_progresso.put({"tipo": "progresso", "atual": i, "total": total})
            
            fila_progresso.put({"tipo": "sucesso", "caminho": diretorio})
        except Exception as e:
            fila_progresso.put({"tipo": "erro", "mensagem": f"Erro ao gerar imagens: {str(e)}"})

    def gerar_zip(self, codigos, caminho_zip, fila_progresso):
        try:
            with zipfile.ZipFile(caminho_zip, "w") as zipf:
                total = len(codigos)
                last_update = time.time()
                
                for i, codigo in enumerate(codigos):
                    img = self.gerar_imagem_qr(codigo)
                    buffer = io.BytesIO()
                    img.save(buffer, format="PNG")
                    buffer.seek(0)
                    
                    nome_arquivo = self._sanitizar_nome(codigo)
                    zipf.writestr(f"{nome_arquivo}.png", buffer.read())

                    # --- MELHORIA: Throttling de Progresso ---
                    should_update, last_update = self._deveria_atualizar_progresso(i, total, last_update)
                    if should_update:
                        fila_progresso.put({"tipo": "progresso", "atual": i, "total": total})
            
            fila_progresso.put({"tipo": "sucesso", "caminho": caminho_zip})
        except Exception as e:
            fila_progresso.put({"tipo": "erro", "mensagem": f"Erro ao gerar ZIP: {str(e)}"})


class ScrollableFrame(ttk.Frame):
    """Frame rolável genérico para Tkinter."""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)


class QRCodeGeneratorUI:
    """
    Classe de Interface (View/Controller).
    Responsável pela renderização da UI, eventos do usuário e orquestração.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de QR Codes Profissional - Otimizado")
        self.root.geometry("1024x860")
        self.root.minsize(950, 600)
        self.root.resizable(True, True)

        # Instancia o Serviço (Lógica separada)
        self.service = QRGeneratorService()

        # Variáveis de Controle da UI
        self.arquivo_fonte = None
        self.preview_data = None
        self.progress_var = tk.DoubleVar()
        self.fila = queue.Queue()
        
        # --- MELHORIA: Fila e Threading para Preview ---
        self.preview_queue = queue.Queue()
        self._preview_gen_count = 0  # Contador para descartar previews antigos (debounce)
        
        self.modo = tk.StringVar(value="texto")
        self.formato_exportacao = tk.StringVar(value="pdf")
        
        # Variáveis de configuração
        self.qr_size = tk.IntVar(value=200)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.logo_path = tk.StringVar(value="")

        self.preview_image_ref = None

        # Traces para atualizar config do serviço e preview
        self._setup_traces()

        # Construção da UI
        self._setup_style()
        self.criar_layout_principal()
        self.criar_conteudo_rolavel()
        self.criar_barra_status_inferior()
        
        self.verificar_fila()
        # Inicia o preview inicial
        self.root.after(100, self.atualizar_preview)

    def _setup_style(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TButton", font=("Helvetica", 10), padding=6)
        self.style.configure("TLabel", font=("Helvetica", 10))
        self.style.configure("Header.TLabel", font=("Helvetica", 14, "bold"))
        self.style.configure("Bold.TLabel", font=("Helvetica", 10, "bold"))

    def _setup_traces(self):
        vars_to_trace = [self.qr_size, self.qr_foreground_color, self.qr_background_color, self.logo_path]
        for var in vars_to_trace:
            var.trace_add("write", lambda *args: self._sync_service_and_preview())

    def _sync_service_and_preview(self):
        """Sincroniza as variáveis da UI com o Serviço e aciona o preview em thread."""
        self.service.atualizar_config(
            self.qr_size.get(),
            self.qr_foreground_color.get(),
            self.qr_background_color.get(),
            self.logo_path.get()
        )
        self.root.after_idle(self.atualizar_preview)

    def criar_layout_principal(self):
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True)

    def criar_conteudo_rolavel(self):
        self.scroll_container = ScrollableFrame(self.main_container)
        self.scroll_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=(20, 0))
        
        content_frame = ttk.Frame(self.scroll_container.scrollable_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            content_frame, text="Gerador de QR Codes Profissional",
            style="Header.TLabel",
        ).pack(pady=(0, 20), anchor="w")

        # --- 1. DADOS ---
        file_frame = ttk.LabelFrame(content_frame, text="1. Seleção de Dados", padding=15)
        file_frame.pack(fill=tk.X, pady=10)
        
        top_file = ttk.Frame(file_frame)
        top_file.pack(fill=tk.X)
        ttk.Button(top_file, text="Selecionar Arquivo (Excel/CSV)", command=self.selecionar_arquivo).pack(side=tk.LEFT)
        self.file_label = ttk.Label(top_file, text="Nenhum arquivo selecionado", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        column_frame = ttk.Frame(file_frame)
        column_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(column_frame, text="Coluna com os dados:", style="Bold.TLabel").pack(side=tk.LEFT)
        self.column_combo = ttk.Combobox(column_frame, state="disabled", width=30)
        self.column_combo.pack(side=tk.LEFT, padx=10)

        # --- 2. CONFIGURAÇÕES ---
        config_frame = ttk.LabelFrame(content_frame, text="2. Configurações do QR Code", padding=15)
        config_frame.pack(fill=tk.X, pady=10)
        
        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X)
        ttk.Radiobutton(mode_frame, text="Modo Texto", variable=self.modo, value="texto", command=self.atualizar_controles_formato).pack(side=tk.LEFT, padx=20)
        ttk.Radiobutton(mode_frame, text="Modo Numérico", variable=self.modo, value="numerico", command=self.atualizar_controles_formato).pack(side=tk.LEFT, padx=20)

        self.texto_controls = ttk.Frame(config_frame)
        ttk.Label(self.texto_controls, text="Máximo de caracteres:").pack(side=tk.LEFT)
        self.max_caracteres = ttk.Spinbox(self.texto_controls, from_=1, to=1000, width=10)
        self.max_caracteres.pack(side=tk.LEFT, padx=5)
        self.max_caracteres.set(250)

        self.numerico_controls = ttk.Frame(config_frame)
        ttk.Label(self.numerico_controls, text="Total de dígitos:").pack(side=tk.LEFT)
        self.total_digitos = ttk.Spinbox(self.numerico_controls, from_=1, to=50, width=10)
        self.total_digitos.pack(side=tk.LEFT, padx=5)
        self.total_digitos.set(10)
        ttk.Label(self.numerico_controls, text="Adicionar número:").pack(side=tk.LEFT, padx=5)
        self.posicao_numero = ttk.Combobox(self.numerico_controls, values=["Antes", "Depois"], width=7, state="readonly")
        self.posicao_numero.pack(side=tk.LEFT, padx=5)
        self.posicao_numero.set("Antes")
        self.numero_adicional = ttk.Entry(self.numerico_controls, width=12)
        self.numero_adicional.pack(side=tk.LEFT, padx=5)
        
        self.atualizar_controles_formato()

        # --- 3. PERSONALIZAÇÃO ---
        custom_frame = ttk.LabelFrame(content_frame, text="3. Personalização", padding=15)
        custom_frame.pack(fill=tk.X, pady=10)

        left_controls = ttk.Frame(custom_frame)
        left_controls.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))

        self._create_input_row(left_controls, "Tamanho (px):", self.qr_size, 100, 1000)
        self._create_color_row(left_controls, "Cor do QR:", self.escolher_cor_qr)
        self._create_color_row(left_controls, "Cor de Fundo:", self.escolher_cor_fundo)
        
        logo_btn_frame = ttk.Frame(left_controls)
        logo_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Label(logo_btn_frame, text="Logotipo:").pack(side=tk.LEFT)
        ttk.Button(logo_btn_frame, text="Selecionar Imagem", command=self.selecionar_logo).pack(side=tk.LEFT, padx=5)

        right_preview_area = ttk.Frame(custom_frame)
        right_preview_area.pack(side=tk.LEFT, fill=tk.Y)

        preview_container = ttk.LabelFrame(right_preview_area, text="Preview (Dados Reais)", padding=10)
        preview_container.pack(side=tk.LEFT, padx=(0, 10))
        
        self.preview_label = ttk.Label(preview_container, text="Sem dados", anchor="center")
        self.preview_label.pack()

        action_container = ttk.Frame(right_preview_area)
        action_container.pack(side=tk.LEFT, fill=tk.Y)
        
        self.generate_button = ttk.Button(
            action_container, text="GERAR\nQR CODE", command=self.iniciar_geracao, state="disabled", width=15
        )
        self.generate_button.pack(expand=True)

        # --- 4. EXPORTAÇÃO ---
        export_frame = ttk.LabelFrame(content_frame, text="4. Exportação", padding=15)
        export_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(export_frame, text="Formato:", style="Bold.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        for formato in ["PDF", "PNG", "SVG", "ZIP"]:
            ttk.Radiobutton(export_frame, text=formato, variable=self.formato_exportacao, value=formato.lower()).pack(side=tk.LEFT, padx=10)

        ttk.Label(content_frame, text="").pack(pady=10)

    def criar_barra_status_inferior(self):
        self.bottom_panel = ttk.Frame(self.main_container, padding=(20, 15, 20, 20), relief="raised")
        self.bottom_panel.pack(side=tk.BOTTOM, fill=tk.X)

        status_area = ttk.Frame(self.bottom_panel)
        status_area.pack(fill=tk.X)

        self.status_label = ttk.Label(status_area, text="Aguardando início...", font=("Helvetica", 9, "italic"))
        self.status_label.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        self.progress_bar = ttk.Progressbar(status_area, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress_bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(
            self.bottom_panel, text="Desenvolvido por Johann Sebastian Dulz | Versão 5.0 (Performance)",
            font=("Helvetica", 8), foreground="#888888",
        ).pack(side=tk.RIGHT)

    # --- Helpers UI ---
    def _create_input_row(self, parent, label_text, variable, min_val, max_val):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=label_text).pack(side=tk.LEFT)
        sb = ttk.Spinbox(row, from_=min_val, to=max_val, textvariable=variable, width=10)
        sb.pack(side=tk.LEFT, padx=5)

    def _create_color_row(self, parent, label_text, command):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=label_text).pack(side=tk.LEFT)
        ttk.Button(row, text="Escolher", command=command, width=10).pack(side=tk.LEFT, padx=5)

    # --- Lógica de Preview e Threading ---
    
    def _gerar_preview_worker(self, generation_id):
        """
        Worker em Thread separada para gerar o preview.
        Verifica se o ID da requisição ainda é válido (não foi substituído por outra mais nova).
        """
        try:
            # Verificação de obsolescência (Debounce)
            if generation_id != self._preview_gen_count:
                return

            # Usa o dado real ou fallback
            data = self.preview_data if self.preview_data else "Preview QR Code"
            
            # Chama o serviço (que usa Cache interno)
            img = self.service.gerar_imagem_qr(data)
            img.thumbnail((180, 180), Image.Resampling.LANCZOS)
            
            # Segunda verificação antes de colocar na fila (para não desenhar preview antigo)
            if generation_id != self._preview_gen_count:
                return
            
            # Converte para PhotoImage na thread de background? 
            # PhotoImage deve ser criado na thread principal do Tkinter.
            # Então passamos o objeto PIL para a fila.
            self.preview_queue.put(("image", img))
            
        except Exception as e:
            # Em caso de erro no preview, não trava o app, apenas ignora
            pass

    def atualizar_preview(self):
        """Aciona o worker de preview em uma nova thread."""
        self._preview_gen_count += 1
        # Passa o ID atual para a thread
        threading.Thread(
            target=self._gerar_preview_worker, 
            args=(self._preview_gen_count,),
            daemon=True
        ).start()

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel/CSV", "*.xlsx;*.csv"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivo_fonte = caminho
            self.file_label.config(text=os.path.basename(caminho), foreground="black")
            try:
                if caminho.endswith(".csv"):
                    df = pd.read_csv(caminho)
                else:
                    df = pd.read_excel(caminho)
                colunas = df.columns.tolist()
                self.column_combo["values"] = colunas
                self.column_combo["state"] = "readonly"
                self.column_combo.set(colunas[0])
                self.generate_button["state"] = "normal"
                
                # Carrega o primeiro dado para o Preview Real
                primeira_coluna = colunas[0]
                primeiro_valor = df[primeira_coluna].dropna().iloc[0] if not df[primeira_coluna].empty else None
                self.preview_data = str(primeiro_valor) if primeiro_valor is not None else None
                
                # Invalida o cache do serviço se os dados mudarem drasticamente
                self.service._cached_qr_object = None
                self.service._cache_data = None
                
                self.atualizar_preview()

            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível carregar o arquivo: {str(e)}")
                self.arquivo_fonte = None
                self.column_combo["state"] = "disabled"
                self.generate_button["state"] = "disabled"

    def escolher_cor_qr(self):
        cor = colorchooser.askcolor(initialcolor=self.qr_foreground_color.get())[1]
        if cor: self.qr_foreground_color.set(cor)

    def escolher_cor_fundo(self):
        cor = colorchooser.askcolor(initialcolor=self.qr_background_color.get())[1]
        if cor: self.qr_background_color.set(cor)

    def selecionar_logo(self):
        caminho = filedialog.askopenfilename(
            filetypes=[("Imagens", "*.png;*.jpg;*.jpeg"), ("Todos os arquivos", "*.*")]
        )
        if caminho: self.logo_path.set(caminho)

    def iniciar_geracao(self):
        coluna_selecionada = self.column_combo.get()
        if not self.arquivo_fonte or not coluna_selecionada:
            messagebox.showwarning("Aviso", "Selecione um arquivo e uma coluna válida.")
            return
        self.alterar_estado_interface(False)
        threading.Thread(target=self.processar_dados, args=(coluna_selecionada,)).start()

    def processar_dados(self, coluna_selecionada):
        """Lê dados e delega a geração para o Serviço."""
        try:
            if self.arquivo_fonte.endswith(".csv"):
                df = pd.read_csv(self.arquivo_fonte)
            else:
                df = pd.read_excel(self.arquivo_fonte)
            
            codigos = df[coluna_selecionada].dropna().astype(str).tolist()
            formato = self.formato_exportacao.get()
            
            # Delegação para o Serviço
            if formato in ["png", "svg"]:
                diretorio = filedialog.askdirectory(title="Selecione a pasta para salvar as imagens")
                if diretorio:
                    self.service.gerar_imagens(codigos, formato, diretorio, self.fila)
                else:
                    self.alterar_estado_interface(True)

            elif formato == "pdf":
                caminho = filedialog.asksaveasfilename(
                    defaultextension=".pdf", filetypes=[("Arquivos PDF", "*.pdf")], title="Salvar arquivo PDF"
                )
                if caminho:
                    self.service.gerar_pdf(codigos, caminho, self.fila)
                else:
                    self.alterar_estado_interface(True)

            elif formato == "zip":
                caminho = filedialog.asksaveasfilename(
                    defaultextension=".zip", filetypes=[("Arquivos ZIP", "*.zip")], title="Salvar arquivo ZIP"
                )
                if caminho:
                    self.service.gerar_zip(codigos, caminho, self.fila)
                else:
                    self.alterar_estado_interface(True)

        except Exception as e:
            self.fila.put({"tipo": "erro", "mensagem": str(e)})

    # --- Gestão de Estado e Fila ---
    def verificar_fila(self):
        """Verifica tanto a fila de progresso principal quanto a fila de preview."""
        # Processa Fila de Progresso
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg["tipo"] == "progresso":
                    self.atualizar_progresso(msg["atual"], msg["total"])
                elif msg["tipo"] == "erro":
                    self.mostrar_erro(msg["mensagem"])
                elif msg["tipo"] == "sucesso":
                    self.mostrar_sucesso(msg["caminho"])
        except queue.Empty:
            pass

        # Processa Fila de Preview
        try:
            # Usa get_nowait para não bloquear
            msg_type, content = self.preview_queue.get_nowait()
            if msg_type == "image":
                # Atualiza a imagem na thread principal (Tkinter requirement)
                self.preview_image_ref = ImageTk.PhotoImage(content)
                self.preview_label.config(image=self.preview_image_ref, text="")
        except queue.Empty:
            pass
        
        self.root.after(50, self.verificar_fila) # Loop mais rápido para reagir ao preview

    def atualizar_progresso(self, atual, total):
        progresso = (atual + 1) / total * 100
        self.progress_var.set(progresso)
        self.status_label.config(text=f"Processando: {atual + 1}/{total} ({int(progresso)}%)")

    def mostrar_erro(self, mensagem):
        messagebox.showerror("Erro na Execução", mensagem)
        self.alterar_estado_interface(True)

    def mostrar_sucesso(self, caminho):
        self.progress_var.set(0)
        self.status_label.config(text="Processo concluído com sucesso!")
        messagebox.showinfo("Sucesso", f"QR Codes gerados com sucesso!\n\nLocal: {caminho}")
        self.alterar_estado_interface(True)

    def alterar_estado_interface(self, habilitar):
        estado = "normal" if habilitar else "disabled"
        self.generate_button["state"] = estado
        self.column_combo["state"] = "readonly" if habilitar else "disabled"
        self.root.config(cursor="watch" if not habilitar else "")
        
        try:
            if self.texto_controls.winfo_children():
                self.texto_controls.winfo_children()[1]["state"] = estado
            if self.numerico_controls.winfo_children():
                self.numerico_controls.winfo_children()[1]["state"] = estado
            self.posicao_numero["state"] = "readonly" if habilitar else "disabled"
            self.numero_adicional["state"] = estado
        except Exception:
            pass

    def atualizar_controles_formato(self):
        if self.modo.get() == "texto":
            self.texto_controls.pack(fill=tk.X, pady=5)
            self.numerico_controls.pack_forget()
        else:
            self.texto_controls.pack_forget()
            self.numerico_controls.pack(fill=tk.X, pady=5)


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGeneratorUI(root)
    root.mainloop()