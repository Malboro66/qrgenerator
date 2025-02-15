import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser
import pandas as pd
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import io
import os
import threading
import queue
from PIL import Image, ImageDraw
import zipfile

class QRCodeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de QR Codes Profissional")
        self.root.geometry("1000x850")
        self.root.resizable(False, False)  # Janela não redimensionável

        # Variáveis de controle
        self.arquivo_fonte = None
        self.progress_var = tk.DoubleVar()
        self.fila = queue.Queue()
        self.modo = tk.StringVar(value="texto")
        self.formato_exportacao = tk.StringVar(value="pdf")
        self.qr_size = tk.IntVar(value=200)
        self.qr_foreground_color = tk.StringVar(value="black")
        self.qr_background_color = tk.StringVar(value="white")
        self.logo_path = tk.StringVar(value="")

        # Criar interface gráfica
        self.criar_interface()
        self.verificar_fila()

    def criar_interface(self):
        """Cria a interface gráfica da aplicação."""
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Título
        ttk.Label(
            main_frame,
            text="Gerador de QR Codes em PDF",
            font=("Helvetica", 16, "bold"),
        ).pack(pady=10)

        # Seção de seleção de arquivo
        file_frame = ttk.LabelFrame(main_frame, text="1. Seleção de Dados", padding=15)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Button(
            file_frame,
            text="Selecionar Arquivo (Excel/CSV)",
            command=self.selecionar_arquivo,
        ).pack(side=tk.TOP, pady=5)
        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(side=tk.TOP)

        # Seção de seleção de coluna
        column_frame = ttk.Frame(file_frame)
        column_frame.pack(pady=10)
        ttk.Label(column_frame, text="Coluna com os dados:").pack(side=tk.LEFT)
        self.column_combo = ttk.Combobox(column_frame, state="disabled", width=25)
        self.column_combo.pack(side=tk.LEFT, padx=10)

        # Configurações do QR Code
        config_frame = ttk.LabelFrame(main_frame, text="2. Configurações do QR Code", padding=15)
        config_frame.pack(fill=tk.X, pady=10)
        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(
            mode_frame,
            text="Modo Texto",
            variable=self.modo,
            value="texto",
            command=self.atualizar_controles_formato,
        ).pack(side=tk.LEFT, padx=15)
        ttk.Radiobutton(
            mode_frame,
            text="Modo Numérico",
            variable=self.modo,
            value="numerico",
            command=self.atualizar_controles_formato,
        ).pack(side=tk.LEFT, padx=15)

        # Controles para modo texto
        self.texto_controls = ttk.Frame(config_frame)
        ttk.Label(self.texto_controls, text="Máximo de caracteres:").pack(side=tk.LEFT)
        self.max_caracteres = ttk.Spinbox(self.texto_controls, from_=1, to=1000, width=8)
        self.max_caracteres.pack(side=tk.LEFT, padx=5)
        self.max_caracteres.set(250)

        # Controles para modo numérico
        self.numerico_controls = ttk.Frame(config_frame)
        ttk.Label(self.numerico_controls, text="Total de dígitos:").pack(side=tk.LEFT)
        self.total_digitos = ttk.Spinbox(self.numerico_controls, from_=1, to=50, width=8)
        self.total_digitos.pack(side=tk.LEFT, padx=5)
        self.total_digitos.set(10)
        ttk.Label(self.numerico_controls, text="Adicionar número:").pack(side=tk.LEFT, padx=5)
        self.posicao_numero = ttk.Combobox(
            self.numerico_controls,
            values=["Antes", "Depois"],
            width=7,
            state="readonly",
        )
        self.posicao_numero.pack(side=tk.LEFT, padx=5)
        self.posicao_numero.set("Antes")
        self.numero_adicional = ttk.Entry(self.numerico_controls, width=12)
        self.numero_adicional.pack(side=tk.LEFT, padx=5)
        self.atualizar_controles_formato()

        # Personalização dos QR Codes
        custom_frame = ttk.LabelFrame(main_frame, text="3. Personalização dos QR Codes", padding=15)
        custom_frame.pack(fill=tk.X, pady=10)

        # Tamanho do QR Code
        size_frame = ttk.Frame(custom_frame)
        size_frame.pack(fill=tk.X, pady=5)
        ttk.Label(size_frame, text="Tamanho do QR Code (pixels):").pack(side=tk.LEFT, padx=5)
        ttk.Spinbox(
            size_frame, from_=100, to=1000, increment=10, textvariable=self.qr_size, width=8
        ).pack(side=tk.LEFT, padx=5)

        # Cor do QR Code
        color_frame = ttk.Frame(custom_frame)
        color_frame.pack(fill=tk.X, pady=5)
        ttk.Label(color_frame, text="Cor do QR Code:").pack(side=tk.LEFT, padx=5)
        ttk.Button(color_frame, text="Escolher Cor", command=self.escolher_cor_qr).pack(
            side=tk.LEFT, padx=5
        )

        # Cor de Fundo
        bg_color_frame = ttk.Frame(custom_frame)
        bg_color_frame.pack(fill=tk.X, pady=5)
        ttk.Label(bg_color_frame, text="Cor de Fundo:").pack(side=tk.LEFT, padx=5)
        ttk.Button(bg_color_frame, text="Escolher Cor", command=self.escolher_cor_fundo).pack(
            side=tk.LEFT, padx=5
        )

        # Logotipo
        logo_frame = ttk.Frame(custom_frame)
        logo_frame.pack(fill=tk.X, pady=5)
        ttk.Label(logo_frame, text="Logotipo:").pack(side=tk.LEFT, padx=5)
        ttk.Button(logo_frame, text="Selecionar Imagem", command=self.selecionar_logo).pack(
            side=tk.LEFT, padx=5
        )

        # Opções de exportação
        export_frame = ttk.LabelFrame(main_frame, text="4. Opções de Exportação", padding=15)
        export_frame.pack(fill=tk.X, pady=10)
        ttk.Label(export_frame, text="Formato de exportação:").pack(side=tk.LEFT)
        for formato in ["PDF", "PNG", "SVG", "ZIP"]:
            ttk.Radiobutton(
                export_frame,
                text=formato,
                variable=self.formato_exportacao,
                value=formato.lower(),
            ).pack(side=tk.LEFT, padx=10)

        # Progresso
        progress_frame = ttk.LabelFrame(main_frame, text="5. Progresso", padding=15)
        progress_frame.pack(fill=tk.X, pady=10)
        self.progress_bar = ttk.Progressbar(
            progress_frame, variable=self.progress_var, maximum=100, mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X)
        self.status_label = ttk.Label(progress_frame, text="Aguardando início...")
        self.status_label.pack(pady=5)

        # Botão de geração
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        self.generate_button = ttk.Button(
            btn_frame, text="Gerar QR Codes", command=self.iniciar_geracao, state="disabled"
        )
        self.generate_button.pack()

        # Rodapé
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        ttk.Label(
            footer_frame,
            text="Desenvolvido por Johann Sebastian Dulz | Versão 2.0",
            font=("Helvetica", 8),
            foreground="#666666",
        ).pack(side=tk.RIGHT)

    def selecionar_arquivo(self):
        """Permite ao usuário selecionar um arquivo Excel ou CSV."""
        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel/CSV", "*.xlsx;*.csv"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivo_fonte = caminho
            self.file_label.config(text=os.path.basename(caminho))
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
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível carregar o arquivo: {str(e)}")
                self.arquivo_fonte = None
                self.column_combo["state"] = "disabled"
                self.generate_button["state"] = "disabled"

    def escolher_cor_qr(self):
        """Permite ao usuário escolher a cor do QR Code."""
        cor = colorchooser.askcolor(initialcolor=self.qr_foreground_color.get())[1]
        if cor:
            self.qr_foreground_color.set(cor)

    def escolher_cor_fundo(self):
        """Permite ao usuário escolher a cor de fundo do QR Code."""
        cor = colorchooser.askcolor(initialcolor=self.qr_background_color.get())[1]
        if cor:
            self.qr_background_color.set(cor)

    def selecionar_logo(self):
        """Permite ao usuário selecionar um logotipo para incluir no QR Code."""
        caminho = filedialog.askopenfilename(
            filetypes=[("Imagens", "*.png;*.jpg;*.jpeg"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.logo_path.set(caminho)

    def gerar_qr_code_personalizado(self, codigo):
        """Gera um QR Code personalizado com tamanho, cores e logotipo."""
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(codigo)
        qr.make(fit=True)
        img = qr.make_image(
            fill_color=self.qr_foreground_color.get(),
            back_color=self.qr_background_color.get(),
        )
        img = img.resize((self.qr_size.get(), self.qr_size.get()), Image.Resampling.LANCZOS)
        if self.logo_path.get():
            try:
                logo = Image.open(self.logo_path.get())
                logo_size = int(self.qr_size.get() * 0.2)
                logo = logo.resize((logo_size, logo_size), Image.Resampling.LANCZOS)
                pos = ((img.size[0] - logo.size[0]) // 2, (img.size[1] - logo.size[1]) // 2)
                img.paste(logo, pos, logo if logo.mode == "RGBA" else None)
            except Exception as e:
                messagebox.showwarning("Aviso", f"Erro ao adicionar logotipo: {str(e)}")
        return img

    def gerar_pdf(self, codigos, caminho_pdf):
        """Gera um PDF com QR Codes personalizados."""
        pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
        largura, altura = A4
        margem = 1 * cm
        colunas, linhas = 3, 4
        qr_por_pagina = colunas * linhas
        for indice, codigo in enumerate(codigos):
            self.fila.put({"tipo": "progresso", "atual": indice, "total": len(codigos)})
            img = self.gerar_qr_code_personalizado(codigo)
            buffer_img = io.BytesIO()
            img.save(buffer_img, format="PNG")
            buffer_img.seek(0)
            pagina, linha, coluna = indice % qr_por_pagina, (indice % qr_por_pagina) // colunas, indice % colunas
            x = margem + coluna * (self.qr_size.get() / 30 + margem)
            y = altura - (margem + (linha + 1) * (self.qr_size.get() / 30 + margem))
            pdf.drawImage(
                ImageReader(buffer_img),
                x,
                y,
                width=self.qr_size.get() / 30,
                height=self.qr_size.get() / 30,
            )
            texto_largura = pdf.stringWidth(codigo, "Helvetica", 8)
            pdf.drawString(x + (self.qr_size.get() / 30 - texto_largura) / 2, y - 12, codigo)
            if (indice + 1) % qr_por_pagina == 0 and indice < len(codigos) - 1:
                pdf.showPage()
        pdf.save()
        self.fila.put({"tipo": "sucesso", "caminho": caminho_pdf})

    def gerar_imagens(self, codigos, formato, diretorio):
        """Gera imagens QR Code em formatos como PNG ou SVG."""
        for codigo in codigos:
            img = self.gerar_qr_code_personalizado(codigo)
            caminho_arquivo = os.path.join(diretorio, f"{codigo}.{formato}")
            if formato == "png":
                img.save(caminho_arquivo, format="PNG")
            elif formato == "svg":
                # Implemente a lógica para salvar como SVG (você pode usar uma biblioteca como svgwrite)
                pass

    def gerar_qr_codes(self, codigos, formato, diretorio):
        """Combina a geração de QR Codes com a exportação."""
        if formato == "png":
            self.gerar_imagens(codigos, "png", diretorio)
        elif formato == "pdf":
            caminho_pdf = os.path.join(diretorio, "qrcodes.pdf")
            self.gerar_pdf(codigos, caminho_pdf)
        elif formato == "zip":
            self.gerar_zip(codigos, diretorio)

    def gerar_zip(self, codigos, diretorio):
        """Cria um arquivo ZIP contendo os QR Codes."""
        caminho_zip = os.path.join(diretorio, "qrcodes.zip")
        with zipfile.ZipFile(caminho_zip, "w") as zipf:
            for codigo in codigos:
                img = self.gerar_qr_code_personalizado(codigo)
                buffer = io.BytesIO()
                img.save(buffer, format="PNG")
                buffer.seek(0)
                zipf.writestr(f"{codigo}.png", buffer.read())

    def verificar_fila(self):
        """Verifica a fila de mensagens para atualizar a interface."""
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
        self.root.after(100, self.verificar_fila)

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso."""
        progresso = (atual + 1) / total * 100
        self.progress_var.set(progresso)
        self.status_label.config(text=f"Processando: {atual + 1}/{total} QR Codes")

    def mostrar_erro(self, mensagem):
        """Exibe uma mensagem de erro."""
        messagebox.showerror("Erro na Execução", mensagem)
        self.alterar_estado_interface(True)

    def mostrar_sucesso(self, caminho):
        """Exibe uma mensagem de sucesso."""
        self.progress_var.set(0)
        self.status_label.config(text="Processo concluído com sucesso!")
        messagebox.showinfo("Sucesso", f"QR Codes gerados com sucesso!\n\nLocal: {caminho}")
        self.alterar_estado_interface(True)

    def alterar_estado_interface(self, habilitar):
        """Habilita ou desabilita elementos da interface."""
        estado = "normal" if habilitar else "disabled"
        self.generate_button["state"] = estado
        self.column_combo["state"] = "readonly" if habilitar else "disabled"
        self.texto_controls.winfo_children()[1]["state"] = estado
        self.numerico_controls.winfo_children()[1]["state"] = estado
        self.posicao_numero["state"] = "readonly" if habilitar else "disabled"
        self.numero_adicional["state"] = estado

    def iniciar_geracao(self):
        """Inicia a geração dos QR Codes em segundo plano."""
        coluna_selecionada = self.column_combo.get()
        if not self.arquivo_fonte or not coluna_selecionada:
            messagebox.showwarning("Aviso", "Selecione um arquivo e uma coluna válida.")
            return
        self.alterar_estado_interface(False)
        threading.Thread(target=self.processar_dados, args=(coluna_selecionada,)).start()

    def processar_dados(self, coluna_selecionada):
        """Processa os dados do arquivo e gera os QR Codes."""
        try:
            if self.arquivo_fonte.endswith(".csv"):
                df = pd.read_csv(self.arquivo_fonte)
            else:
                df = pd.read_excel(self.arquivo_fonte)
            codigos = df[coluna_selecionada].dropna().astype(str).tolist()
            caminho_saida = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            )
            if caminho_saida:
                self.gerar_pdf(codigos, caminho_saida)
        except Exception as e:
            self.fila.put({"tipo": "erro", "mensagem": str(e)})

    def atualizar_controles_formato(self):
        """Atualiza os controles visíveis com base no modo selecionado."""
        if self.modo.get() == "texto":
            self.texto_controls.pack(fill=tk.X, pady=5)
            self.numerico_controls.pack_forget()
        else:
            self.texto_controls.pack_forget()
            self.numerico_controls.pack(fill=tk.X, pady=5)


if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()