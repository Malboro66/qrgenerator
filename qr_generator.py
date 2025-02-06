# Importações necessárias
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
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

class QRCodeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de QR Codes Profissional")
        self.root.geometry("750x550")
        self.root.resizable(False, False)  # Janela não redimensionável
        
        self.arquivo_fonte = None
        self.progress_var = tk.DoubleVar()
        self.fila = queue.Queue()
        self.modo = tk.StringVar(value='texto')
        
        self.criar_interface()
        self.verificar_fila()

    def criar_interface(self):
        # Configuração do layout principal
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Título centralizado
        ttk.Label(
            main_frame,
            text="Gerador de QR Codes em PDF",
            font=('Helvetica', 16, 'bold')
        ).pack(pady=10)

        # Seção de seleção de arquivo
        file_frame = ttk.LabelFrame(main_frame, text="1. Seleção de Dados", padding=15)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            file_frame,
            text="Selecionar Arquivo (Excel/CSV)",
            command=self.selecionar_arquivo
        ).pack(side=tk.TOP, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(side=tk.TOP)

        # Seção de seleção de coluna
        column_frame = ttk.Frame(file_frame)
        column_frame.pack(pady=10)
        ttk.Label(column_frame, text="Coluna com os dados:").pack(side=tk.LEFT)
        self.column_combo = ttk.Combobox(column_frame, state="disabled", width=25)
        self.column_combo.pack(side=tk.LEFT, padx=10)

        # Seção de configurações
        config_frame = ttk.LabelFrame(main_frame, text="2. Configurações do QR Code", padding=15)
        config_frame.pack(fill=tk.X, pady=10)

        # Controles de modo
        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(
            mode_frame,
            text="Modo Texto",
            variable=self.modo,
            value='texto',
            command=self.atualizar_controles_formato
        ).pack(side=tk.LEFT, padx=15)
        
        ttk.Radiobutton(
            mode_frame,
            text="Modo Numérico",
            variable=self.modo,
            value='numerico',
            command=self.atualizar_controles_formato
        ).pack(side=tk.LEFT, padx=15)

        # Controles para texto
        self.texto_controls = ttk.Frame(config_frame)
        ttk.Label(self.texto_controls, text="Máximo de caracteres:").pack(side=tk.LEFT)
        self.max_caracteres = ttk.Spinbox(self.texto_controls, from_=1, to=1000, width=8)
        self.max_caracteres.pack(side=tk.LEFT, padx=5)
        self.max_caracteres.set(250)

        # Controles para numérico
        self.numerico_controls = ttk.Frame(config_frame)
        ttk.Label(self.numerico_controls, text="Total de dígitos:").pack(side=tk.LEFT)
        self.total_digitos = ttk.Spinbox(self.numerico_controls, from_=1, to=50, width=8)
        self.total_digitos.pack(side=tk.LEFT, padx=5)
        self.total_digitos.set(10)
        
        ttk.Label(self.numerico_controls, text="Adicionar número:").pack(side=tk.LEFT, padx=5)
        self.posicao_numero = ttk.Combobox(
            self.numerico_controls,
            values=['Antes', 'Depois'],
            width=7,
            state='readonly'
        )
        self.posicao_numero.pack(side=tk.LEFT, padx=5)
        self.posicao_numero.set('Antes')
        
        self.numero_adicional = ttk.Entry(self.numerico_controls, width=12)
        self.numero_adicional.pack(side=tk.LEFT, padx=5)
        
        self.atualizar_controles_formato()

        # Seção de progresso
        progress_frame = ttk.LabelFrame(main_frame, text="3. Progresso", padding=15)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X)
        
        self.status_label = ttk.Label(progress_frame, text="Aguardando início...")
        self.status_label.pack(pady=5)

        # Botão de ação principal
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        self.generate_button = ttk.Button(
            btn_frame,
            text="Gerar QR Codes em PDF",
            command=self.iniciar_geracao,
            state="disabled"
        )
        self.generate_button.pack()

        # Rodapé
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(
            footer_frame,
            text="Desenvolvido por Johann Sebastian Dulz | Versão 2.0",
            font=('Helvetica', 8),
            foreground="#666666"
        ).pack(side=tk.RIGHT)

    def atualizar_controles_formato(self):
        if self.modo.get() == 'texto':
            self.texto_controls.pack(fill=tk.X, pady=5)
            self.numerico_controls.pack_forget()
        else:
            self.numerico_controls.pack(fill=tk.X, pady=5)
            self.texto_controls.pack_forget()

    def verificar_extensao(self, arquivo):
        extensoes_validas = ('.xlsx', '.xls', '.csv')
        return os.path.splitext(arquivo)[1].lower() in extensoes_validas

    def determinar_tipo_arquivo(self, arquivo):
        extensao = os.path.splitext(arquivo)[1].lower()
        return 'excel' if extensao in ('.xlsx', '.xls') else 'csv'

    def ler_arquivo(self, caminho):
        try:
            if self.tipo_arquivo == 'excel':
                return pd.read_excel(caminho)
            elif self.tipo_arquivo == 'csv':
                try:
                    return pd.read_csv(caminho, encoding='utf-8', delimiter=',')
                except UnicodeDecodeError:
                    return pd.read_csv(caminho, encoding='latin-1', delimiter=';')
        except Exception as e:
            raise ValueError(f"Erro na leitura do arquivo: {str(e)}")

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            filetypes=[("Arquivos Suportados", "*.xlsx;*.xls;*.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if not arquivo:
            return
            
        if not self.verificar_extensao(arquivo):
            messagebox.showerror("Erro", "Formato de arquivo não suportado!")
            return
            
        self.arquivo_fonte = arquivo
        self.tipo_arquivo = self.determinar_tipo_arquivo(arquivo)
        self.file_label.config(text=os.path.basename(arquivo))
        
        try:
            df = self.ler_arquivo(arquivo)
            colunas = list(df.columns)
            self.column_combo['values'] = colunas
            self.column_combo['state'] = 'readonly'
            self.column_combo.set(colunas[0])
            self.generate_button['state'] = 'normal'
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler arquivo:\n{str(e)}")
            self.limpar_selecao()

    def limpar_selecao(self):
        self.arquivo_fonte = None
        self.file_label.config(text="Nenhum arquivo selecionado")
        self.column_combo.set('')
        self.column_combo['state'] = 'disabled'
        self.generate_button['state'] = 'disabled'

    def validar_formato(self):
        try:
            if self.modo.get() == 'numerico':
                if not self.numero_adicional.get().isdigit():
                    raise ValueError("O número adicional deve conter apenas dígitos")
                
                total = int(self.total_digitos.get())
                add_len = len(self.numero_adicional.get())
                
                if add_len >= total:
                    raise ValueError("O número adicional não pode ser maior que o total de dígitos")
                
            return True
        except ValueError as e:
            self.fila.put({'tipo': 'erro', 'mensagem': str(e)})
            return False

    def processar_codigos(self, codigos):
        if self.modo.get() == 'texto':
            max_chars = int(self.max_caracteres.get())
            return [str(c)[:max_chars] for c in codigos]
        else:
            total_digitos = int(self.total_digitos.get())
            numero_add = self.numero_adicional.get()
            add_len = len(numero_add)
            code_len = total_digitos - add_len
            posicao = self.posicao_numero.get().lower()
            
            processed = []
            for codigo in codigos:
                codigo_str = str(codigo).strip()
                if not codigo_str.isdigit():
                    raise ValueError(f"Dado não numérico encontrado: {codigo_str}")
                
                codigo_ajustado = codigo_str.zfill(code_len)[-code_len:]
                
                if posicao == 'antes':
                    novo_codigo = f"{numero_add}{codigo_ajustado}"
                else:
                    novo_codigo = f"{codigo_ajustado}{numero_add}"
                
                processed.append(novo_codigo[:total_digitos])
            
            return processed

    def validar_dados(self):
        try:
            df = self.ler_arquivo(self.arquivo_fonte)
            coluna = self.column_combo.get()
            
            if coluna not in df.columns:
                raise ValueError(f"Coluna '{coluna}' não encontrada no arquivo")
                
            codigos = df[coluna].astype(str).str.strip().tolist()
            
            if self.modo.get() == 'numerico':
                if not all(c.isdigit() for c in codigos if c != ''):
                    raise ValueError("Modo numérico selecionado, mas a coluna contém dados não numéricos")
            
            codigos = [c for c in codigos if c != '']
            
            if not codigos:
                raise ValueError("Nenhum dado válido encontrado na coluna selecionada")
                
            return codigos
            
        except Exception as e:
            self.fila.put({'tipo': 'erro', 'mensagem': str(e)})
            return None

    def iniciar_geracao(self):
        if not self.validar_formato():
            return
            
        caminho_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")]
        )
        
        if not caminho_pdf:
            return
            
        codigos = self.validar_dados()
        if not codigos:
            return
            
        try:
            codigos_processados = self.processar_codigos(codigos)
        except Exception as e:
            self.fila.put({'tipo': 'erro', 'mensagem': str(e)})
            return
            
        self.alterar_estado_interface(False)
        
        thread = threading.Thread(
            target=self.gerar_pdf,
            args=(codigos_processados, caminho_pdf),
            daemon=True
        )
        thread.start()

    def gerar_pdf(self, codigos, caminho_pdf):
        try:
            pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
            largura, altura = A4
            
            tamanho_qr = 5 * cm
            margem = 1 * cm
            colunas = 3
            linhas = 4
            qr_por_pagina = colunas * linhas
            
            config_qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=10,
                border=4,
            )
            
            for indice, codigo in enumerate(codigos):
                self.fila.put({
                    'tipo': 'progresso',
                    'atual': indice,
                    'total': len(codigos)
                })
                
                config_qr.clear()
                config_qr.add_data(codigo)
                config_qr.make(fit=True)
                
                img = config_qr.make_image(fill_color="black", back_color="white")
                buffer_img = io.BytesIO()
                img.save(buffer_img, format='PNG')
                buffer_img.seek(0)
                
                pagina = indice % qr_por_pagina
                linha = pagina // colunas
                coluna = pagina % colunas
                
                x = margem + coluna * (tamanho_qr + margem)
                y = altura - (margem + (linha + 1) * (tamanho_qr + margem))
                
                pdf.drawImage(
                    ImageReader(buffer_img),
                    x, 
                    y, 
                    width=tamanho_qr, 
                    height=tamanho_qr
                )
                
                pdf.setFont("Helvetica", 8)
                texto_largura = pdf.stringWidth(codigo, "Helvetica", 8)
                pdf.drawString(x + (tamanho_qr - texto_largura)/2, y - 12, codigo)
                
                if (indice + 1) % qr_por_pagina == 0 and indice < len(codigos) - 1:
                    pdf.showPage()

            pdf.save()
            self.fila.put({'tipo': 'sucesso', 'caminho': caminho_pdf})
            
        except Exception as e:
            self.fila.put({'tipo': 'erro', 'mensagem': f"Erro na geração do PDF: {str(e)}"})

    def verificar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg['tipo'] == 'progresso':
                    self.atualizar_progresso(msg['atual'], msg['total'])
                elif msg['tipo'] == 'erro':
                    self.mostrar_erro(msg['mensagem'])
                elif msg['tipo'] == 'sucesso':
                    self.mostrar_sucesso(msg['caminho'])
        except queue.Empty:
            pass
        self.root.after(100, self.verificar_fila)

    def atualizar_progresso(self, atual, total):
        progresso = (atual + 1) / total * 100
        self.progress_var.set(progresso)
        self.status_label.config(text=f"Processando: {atual + 1}/{total} QR Codes")

    def mostrar_erro(self, mensagem):
        messagebox.showerror("Erro na Execução", mensagem)
        self.alterar_estado_interface(True)

    def mostrar_sucesso(self, caminho):
        self.progress_var.set(0)
        self.status_label.config(text="Processo concluído com sucesso!")
        messagebox.showinfo(
            "Sucesso", 
            f"PDF gerado com sucesso!\n\nLocal: {caminho}"
        )
        self.alterar_estado_interface(True)

    def alterar_estado_interface(self, habilitar):
        estados = 'normal' if habilitar else 'disabled'
        self.select_button['state'] = estados
        self.generate_button['state'] = estados
        self.column_combo['state'] = 'readonly' if habilitar else 'disabled'
        self.texto_controls.winfo_children()[1]['state'] = estados
        self.numerico_controls.winfo_children()[1]['state'] = estados
        self.posicao_numero['state'] = 'readonly' if habilitar else 'disabled'
        self.numero_adicional['state'] = estados

if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()