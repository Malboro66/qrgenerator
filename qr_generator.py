# Importações necessárias
import tkinter as tk  # Framework para interface gráfica
from tkinter import filedialog, ttk, messagebox  # Componentes da interface gráfica
import pandas as pd  # Para manipulação de arquivos Excel/CSV
import qrcode  # Para gerar códigos QR
from reportlab.pdfgen import canvas  # Geração de PDF
from reportlab.lib.pagesizes import A4  # Tamanho de página padrão
from reportlab.lib.units import cm  # Unidades métricas
from reportlab.lib.utils import ImageReader  # Para manipular imagens no PDF
import io  # Para manipulação de fluxos de bytes
import os  # Operações no sistema de arquivos
import threading  # Para executar tarefas em segundo plano
import queue  # Para comunicação segura entre threads


class QRCodeGenerator:
    """
    Classe principal do aplicativo para gerar códigos QR a partir de dados Excel/CSV
    e salvá-los em PDF.
    """
    def __init__(self, root):
        """
        Inicializa a janela do aplicativo e configura as configurações básicas.
        
        Args:
            root: A janela principal do tkinter
        """
        self.root = root
        self.root.title("Gerador de QR Codes")
        self.root.geometry("600x400")  # Define o tamanho da janela
        
        # Inicializa variáveis de instância
        self.arquivo_excel = None  # Armazena o caminho do arquivo selecionado
        self.progress_var = tk.DoubleVar()  # Controla o status da barra de progresso
        self.fila = queue.Queue()  # Fila de mensagens thread-safe
        
        # Cria elementos da interface gráfica
        self.criar_widgets()
        # Inicia o monitoramento da fila de mensagens
        self.verificar_fila()

    def criar_widgets(self):
        """
        Cria e organiza todos os elementos da interface gráfica na janela principal.
        Inclui botões, rótulos, barra de progresso e menu suspenso.
        """
        # Frame principal com preenchimento
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Rótulo do título
        title_label = ttk.Label(
            main_frame, 
            text="Gerador de QR Codes em PDF", 
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=20)
        
        # Botão de seleção de arquivo
        self.select_button = ttk.Button(
            main_frame,
            text="Selecionar Arquivo (Excel/CSV)",
            command=self.selecionar_arquivo
        )
        self.select_button.pack(pady=10)
        
        # Rótulo para mostrar o nome do arquivo selecionado
        self.file_label = ttk.Label(main_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(pady=5)
        
        # Frame para o menu suspenso de seleção de coluna
        column_frame = ttk.Frame(main_frame)
        column_frame.pack(pady=10)
        
        ttk.Label(column_frame, text="Selecione a coluna dos códigos:").pack(side=tk.LEFT)
        self.column_combo = ttk.Combobox(column_frame, state="disabled")
        self.column_combo.pack(side=tk.LEFT, padx=5)
        
        # Barra de progresso para status da geração
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=20)
        
        # Botão de geração (inicialmente desativado)
        self.generate_button = ttk.Button(
            main_frame,
            text="Gerar QR Codes",
            command=self.iniciar_geracao,
            state="disabled"
        )
        self.generate_button.pack(pady=10)
        
        # Rótulo de status para mostrar o progresso
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=5)

        # Rodapé com créditos
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
        """
        Verifica se a extensão do arquivo é suportada (.xlsx, .xls, .csv)
        
        Args:
            arquivo: Caminho do arquivo para verificar
            
        Returns:
            bool: True se a extensão é suportada, False caso contrário
        """
        extensoes_validas = ('.xlsx', '.xls', '.csv')
        return os.path.splitext(arquivo)[1].lower() in extensoes_validas

    def determinar_tipo_arquivo(self, arquivo):
        """
        Determina se o arquivo é Excel ou CSV com base na extensão
        
        Args:
            arquivo: Caminho do arquivo
            
        Returns:
            str: 'excel' ou 'csv'
        """
        extensao = os.path.splitext(arquivo)[1].lower()
        return 'excel' if extensao in ('.xlsx', '.xls') else 'csv'

    def ler_arquivo(self, caminho):
        """
        Lê o arquivo de entrada (Excel ou CSV) usando pandas
        
        Args:
            caminho: Caminho do arquivo
            
        Returns:
            pandas.DataFrame: Os dados carregados
            
        Raises:
            ValueError: Se o arquivo não puder ser lido
        """
        try:
            if self.tipo_arquivo == 'excel':
                return pd.read_excel(caminho)
            elif self.tipo_arquivo == 'csv':
                try:
                    return pd.read_csv(caminho, encoding='utf-8', delimiter=',')
                except UnicodeDecodeError:
                    # Tenta codificação alternativa se UTF-8 falhar
                    return pd.read_csv(caminho, encoding='latin-1', delimiter=';')
        except Exception as e:
            raise ValueError(f"Erro na leitura do arquivo: {str(e)}")

    def selecionar_arquivo(self):
        """
        Manipula a seleção de arquivo via diálogo e carrega o conteúdo do arquivo
        Atualiza elementos da interface com base na seleção
        """
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
        """Redefine os elementos da interface para o estado inicial"""
        self.arquivo_fonte = None
        self.file_label.config(text="Nenhum arquivo selecionado")
        self.column_combo.set('')
        self.column_combo['state'] = 'disabled'
        self.generate_button['state'] = 'disabled'

    def verificar_fila(self):
        """
        Verifica a fila de mensagens para atualizações da thread de trabalho
        Atualiza a interface com base nas mensagens recebidas
        """
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
        # Agenda próxima verificação
        self.root.after(100, self.verificar_fila)

    def atualizar_interface_progresso(self, atual, total):
        """Atualiza a barra de progresso e o rótulo de status"""
        progresso = (atual + 1) / total * 100
        self.progress_var.set(progresso)
        self.status_label.config(text=f"Gerando QR Code {atual + 1} de {total}")

    def tratar_erro(self, mensagem):
        """Manipula e exibe mensagens de erro"""
        messagebox.showerror("Erro", mensagem)
        self.redefinir_interface()

    def tratar_sucesso(self, caminho):
        """Manipula a geração bem-sucedida do PDF"""
        self.progress_var.set(0)
        self.status_label.config(text="Concluído!")
        messagebox.showinfo("Sucesso", f"PDF gerado com sucesso!\nLocal: {caminho}")
        self.redefinir_interface()

    def redefinir_interface(self):
        """Redefine os elementos da interface após a conclusão da operação"""
        self.select_button['state'] = 'normal'
        self.generate_button['state'] = 'normal'
        self.column_combo['state'] = 'readonly'

    def validar_dados(self):
        """
        Valida os dados de entrada do arquivo e coluna selecionados
        
        Returns:
            list: Lista de códigos válidos para gerar QR codes
            None: Se a validação falhar
        """
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
        """
        Inicia o processo de geração do código QR
        Abre diálogo de salvamento e inicia thread de trabalho
        """
        caminho_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if not caminho_pdf:
            return
            
        codigos = self.validar_dados()
        if not codigos:
            return
            
        # Desativa elementos da interface durante a geração
        self.select_button['state'] = 'disabled'
        self.generate_button['state'] = 'disabled'
        self.column_combo['state'] = 'disabled'
        
        # Inicia geração em thread de segundo plano
        thread = threading.Thread(
            target=self.gerar_pdf,
            args=(codigos, caminho_pdf),
            daemon=True
        )
        thread.start()

    def gerar_pdf(self, codigos, caminho_pdf):
        """
        Gera PDF com códigos QR
        
        Args:
            codigos: Lista de códigos para gerar QR codes
            caminho_pdf: Caminho do arquivo PDF de saída
        """
        try:
            pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
            largura, altura = A4
            
            # Define parâmetros de layout
            tamanho_qr = 5 * cm
            margem = 1 * cm
            colunas = 3
            linhas = 4
            qr_por_pagina = colunas * linhas
            
            # Configura geração do código QR
            config_qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=10,
                border=4,
            )
            
            # Gera códigos QR e adiciona ao PDF
            for indice, codigo in enumerate(codigos):
                # Atualiza progresso
                self.fila.put({
                    'tipo': 'progresso',
                    'atual': indice,
                    'total': len(codigos)
                })
                
                # Gera código QR
                config_qr.clear()
                config_qr.add_data(codigo)
                config_qr.make(fit=True)
                
                img = config_qr.make_image(fill_color="black", back_color="white")
                buffer_img = io.BytesIO()
                img.save(buffer_img, format='PNG')
                buffer_img.seek(0)
                
                # Calcula posição na página
                pagina = indice % qr_por_pagina
                linha = pagina // colunas
                coluna = pagina % colunas
                
                x = margem + coluna * (tamanho_qr + margem)
                y = altura - (margem + (linha + 1) * (tamanho_qr + margem))
                
                # Adiciona código QR ao PDF
                pdf.drawImage(
                    ImageReader(buffer_img),
                    x, 
                    y, 
                    width=tamanho_qr, 
                    height=tamanho_qr
                )
                
                # Adiciona texto abaixo do código QR
                pdf.setFont("Helvetica", 8)
                texto_largura = pdf.stringWidth(codigo, "Helvetica", 8)
                pdf.drawString(x + (tamanho_qr - texto_largura)/2, y - 12, codigo)
                
                # Inicia nova página se necessário
                if (indice + 1) % qr_por_pagina == 0 and indice < len(codigos) - 1:
                    pdf.showPage()

            pdf.save()
            self.fila.put({'tipo': 'sucesso', 'caminho': caminho_pdf})
            
        except Exception as e:
            self.fila.put({'tipo': 'erro', 'mensagem': f"Erro na geração: {str(e)}"})


# Ponto de entrada do aplicativo
if __name__ == "__main__":
    root = tk.Tk()
    app = QRCodeGenerator(root)
    root.mainloop()