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
        self.root.title("Gerador de QR Codes")
        self.root.geometry("600x400")
        
        self.arquivo_excel = None
        self.progress_var = tk.DoubleVar()
        self.fila = queue.Queue()
        
        self.criar_widgets()
        self.verificar_fila()

    def criar_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(
            main_frame, 
            text="Gerador de QR Codes em PDF", 
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=20)
        
        self.select_button = ttk.Button(
            main_frame,
            text="Selecionar Arquivo (Excel/CSV)",
            command=self.selecionar_arquivo
        )
        self.select_button.pack(pady=10)
        
        self.file_label = ttk.Label(main_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(pady=5)
        
        column_frame = ttk.Frame(main_frame)
        column_frame.pack(pady=10)
        
        ttk.Label(column_frame, text="Selecione a coluna dos códigos:").pack(side=tk.LEFT)
        self.column_combo = ttk.Combobox(column_frame, state="disabled")
        self.column_combo.pack(side=tk.LEFT, padx=5)
        
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=20)
        
        self.generate_button = ttk.Button(
            main_frame,
            text="Gerar QR Codes",
            command=self.iniciar_geracao,
            state="disabled"
        )
        self.generate_button.pack(pady=10)
        
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=5)

        # Footer com créditos
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        credit_label = ttk.Label(
            footer_frame,
            text="Desenvolvido por Johann Sebastian Dulz",
            font=('Helvetica', 8),
            foreground="#666666"
        )
        credit_label.pack(side=tk.RIGHT)

        
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
            messagebox.showerror(
                "Erro de Extensão",
                "Formato não suportado!\nUse .xlsx, .xls ou .csv"
            )
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

    def verificar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg['tipo'] == 'progresso':
                    self.atualizar_interface_progresso(msg['atual'], msg['total'])
                elif msg['tipo'] == 'erro':
                    self.tratar_erro(msg['mensagem'])
                elif msg['tipo'] == 'sucesso':
                    self.tratar_sucesso(msg['caminho'])
        except queue.Empty:
            pass
        self.root.after(100, self.verificar_fila)

    def atualizar_interface_progresso(self, atual, total):
        progresso = (atual + 1) / total * 100
        self.progress_var.set(progresso)
        self.status_label.config(text=f"Gerando QR Code {atual + 1} de {total}")

    def tratar_erro(self, mensagem):
        messagebox.showerror("Erro", mensagem)
        self.redefinir_interface()

    def tratar_sucesso(self, caminho):
        self.progress_var.set(0)
        self.status_label.config(text="Concluído!")
        messagebox.showinfo("Sucesso", f"PDF gerado com sucesso!\nLocal: {caminho}")
        self.redefinir_interface()

    def redefinir_interface(self):
        self.select_button['state'] = 'normal'
        self.generate_button['state'] = 'normal'
        self.column_combo['state'] = 'readonly'

    def validar_dados(self):
        try:
            df = self.ler_arquivo(self.arquivo_fonte)
            coluna = self.column_combo.get()
            
            if coluna not in df.columns:
                raise ValueError(f"Coluna '{coluna}' não encontrada")
                
            codigos = df[coluna].astype(str).str.strip()
            
            if codigos.empty:
                raise ValueError("A coluna selecionada está vazia")
                
            codigos = codigos[codigos != ''].tolist()
            
            if not codigos:
                raise ValueError("Nenhum código válido encontrado na coluna")
                
            return codigos
            
        except Exception as e:
            self.fila.put({'tipo': 'erro', 'mensagem': str(e)})
            return None

    def iniciar_geracao(self):
        caminho_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if not caminho_pdf:
            return
            
        codigos = self.validar_dados()
        if not codigos:
            return
            
        self.select_button['state'] = 'disabled'
        self.generate_button['state'] = 'disabled'
        self.column_combo['state'] = 'disabled'
        
        thread = threading.Thread(
            target=self.gerar_pdf,
            args=(codigos, caminho_pdf),
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
                
                # Correção aplicada aqui usando ImageReader
                pdf.drawImage(
                    ImageReader(buffer_img),  # Conversão do buffer
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
            self.fila.put({'tipo': 'erro', 'mensagem': f"Erro na geração: {str(e)}"})

if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()