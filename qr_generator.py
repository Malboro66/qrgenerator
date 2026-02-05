import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import threading
import queue

class BuscadorService:
    """Responsável pela lógica de busca e carregamento de arquivos (Background)."""
    def __init__(self, fila_resultados):
        self.fila = fila_resultados

    def buscar_e_carregar(self, diretorio_raiz, termo_busca):
        """Procura a pasta e carrega as imagens em background."""
        # 1. Busca
        caminho_pasta = None
        termo_lower = termo_busca.lower().strip()
        
        if not os.path.exists(diretorio_raiz):
            self.fila.put({"status": "error", "msg": "Diretório raiz não encontrado."})
            return

        # Busca exata (pode-se adicionar lógica de busca parcial aqui)
        try:
            for pasta in os.listdir(diretorio_raiz):
                caminho_completo = os.path.join(diretorio_raiz, pasta)
                if os.path.isdir(caminho_completo) and pasta.lower() == termo_lower:
                    caminho_pasta = caminho_completo
                    nome_peca = pasta
                    break
        except Exception as e:
            self.fila.put({"status": "error", "msg": str(e)})
            return

        if not caminho_pasta:
            self.fila.put({"status": "not_found"})
            return

        # 2. Carregamento de Imagens
        try:
            arquivos = os.listdir(caminho_pasta)
            extensoes = ('.jpg', '.jpeg', '.png', '.bmp', '.gif')
            imagens = [f for f in arquivos if f.lower().endswith(extensoes)]
            
            total = len(imagens)
            if total == 0:
                self.fila.put({"status": "no_images"})
                return

            self.fila.put({"status": "start", "total": total, "nome": nome_peca})

            for i, arquivo in enumerate(imagens):
                caminho_img = os.path.join(caminho_pasta, arquivo)
                
                # Carrega e redimensiona
                try:
                    img = Image.open(caminho_img)
                    img.thumbnail((250, 250)) # Tamanho um pouco maior para melhor visualização
                    
                    # Converte para PhotoImage
                    photo = ImageTk.PhotoImage(img)
                    
                    # Envia para a UI thread
                    self.fila.put({
                        "status": "progress",
                        "data": (arquivo, photo),
                        "current": i,
                        "total": total
                    })
                except Exception as e:
                    print(f"Erro ao carregar {arquivo}: {e}")
            
            self.fila.put({"status": "done"})

        except Exception as e:
            self.fila.put({"status": "error", "msg": str(e)})

class VisualizadorPecas:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Peças Pro - Inventory Photo Manager")
        self.root.geometry("1100x800")
        self.root.minsize(800, 600)
        
        # Configuração de Estilo (Tema)
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("TLabel", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))

        # Variáveis de Controle
        self.diretorio_raiz = tk.StringVar()
        self.pesquisa_var = tk.StringVar()
        self.imagens_ativas = [] # Evita garbage collection
        self.worker_thread = None
        self.fila = queue.Queue()
        self.grid_row = 0
        self.grid_col = 0
        self.max_cols = 3 # Número de imagens por linha

        # Interface
        self.criar_interface()
        
        # Inicia o loop de verificação da fila
        self.verificar_fila()

    def criar_interface(self):
        # --- Topo: Controles ---
        controle_frame = ttk.Frame(self.root, padding=10)
        controle_frame.pack(fill="x")

        # Linha 1: Pasta e Busca
        linha1 = ttk.Frame(controle_frame)
        linha1.pack(fill="x", pady=5)

        ttk.Button(linha1, text="📂 Selecionar Pasta Raiz", command=self.selecionar_pasta).pack(side="left", padx=5)
        
        ttk.Label(linha1, text="Código da Peça:").pack(side="left", padx=(20, 5))
        entrada = ttk.Entry(linha1, textvariable=self.pesquisa_var, width=20, font=("Arial", 10))
        entrada.pack(side="left", padx=5)
        entrada.bind("<Return>", lambda e: self.iniciar_busca())
        
        ttk.Button(linha1, text="🔍 Buscar", command=self.iniciar_busca).pack(side="left", padx=5)

        # Barra de Progresso (Invisível inicialmente)
        self.progress_frame = ttk.Frame(controle_frame)
        self.progress_frame.pack(fill="x", pady=5)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="indeterminate")
        self.lbl_progresso = ttk.Label(self.progress_frame, text="")

        # --- Meio: Área de Scroll ---
        container_scroll = ttk.Frame(self.root)
        container_scroll.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(container_scroll, bg="white")
        self.scrollbar = ttk.Scrollbar(container_scroll, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = ttk.Frame(self.canvas, bg="white")
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # --- Rodapé: Status ---
        self.status_var = tk.StringVar(value="Aguardando seleção de pasta...")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta onde estão as pastas das peças")
        if pasta:
            self.diretorio_raiz.set(pasta)
            self.status_var.set(f"Pasta selecionada: {pasta}")
            messagebox.showinfo("Pasta Selecionada", "Pasta raiz definida com sucesso.\nAgora digite o código da peça para buscar.")

    def iniciar_busca(self):
        termo = self.pesquisa_var.get().strip()
        if not termo:
            messagebox.showwarning("Aviso", "Digite um código para buscar.")
            return
        
        if not self.diretorio_raiz.get():
            messagebox.showwarning("Aviso", "Selecione a pasta raiz primeiro.")
            return

        # Limpa a UI anterior
        self.limpar_visualizacao()
        
        # Inicia Thread
        self.worker_thread = threading.Thread(
            target=self._worker_target, 
            args=(termo,), 
            daemon=True
        )
        self.worker_thread.start()

        # Atualiza UI para estado de carregamento
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.progress_bar.start(10)
        self.lbl_progresso.pack(side="left", padx=5)
        self.lbl_progresso.config(text=f"Buscando por '{termo}'...")
        self.status_var.set("Carregando imagens... (Interface responsiva)")

    def _worker_target(self, termo):
        """Função executada na thread separada."""
        service = BuscadorService(self.fila)
        service.buscar_e_carregar(self.diretorio_raiz.get(), termo)

    def limpar_visualizacao(self):
        """Destroi todos os widgets do grid e limpa memória."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.imagens_ativas.clear()
        self.grid_row = 0
        self.grid_col = 0

    def verificar_fila(self):
        """Verifica mensagens da thread de background e atualiza a UI na thread principal."""
        try:
            msg = self.fila.get_nowait()
            
            if msg["status"] == "start":
                # Exibe título
                ttk.Label(
                    self.scrollable_frame, 
                    text=f"Resultados para: {msg['nome']}", 
                    style="Header.TLabel",
                    background="white"
                ).pack(pady=10)
                self.progress_bar.config(mode="determinate", maximum=msg["total"])
                self.progress_bar['value'] = 0

            elif msg["status"] == "progress":
                # Adiciona imagem ao grid
                nome_arquivo, photo = msg["data"]
                self.adicionar_imagem_grid(nome_arquivo, photo)
                self.imagens_ativas.append(photo)
                
                # Atualiza barra
                self.progress_bar['value'] = msg["current"] + 1
                self.lbl_progresso.config(text=f"Carregando: {msg['current']+1}/{msg['total']}")

            elif msg["status"] == "done":
                self.finalizar_carregamento("Imagens carregadas com sucesso!")
            
            elif msg["status"] == "not_found":
                ttk.Label(self.scrollable_frame, text="Peça não encontrada.", font=("Arial", 14), foreground="red", background="white").pack(pady=20)
                self.finalizar_carregamento("Peça não encontrada no índice.")
            
            elif msg["status"] == "no_images":
                ttk.Label(self.scrollable_frame, text="Pasta encontrada, mas sem imagens.", background="white").pack(pady=20)
                self.finalizar_carregamento("Pasta vazia.")

            elif msg["status"] == "error":
                messagebox.showerror("Erro", msg["msg"])
                self.finalizar_carregamento("Erro ao processar.")

        except queue.Empty:
            pass
        
        self.root.after(50, self.verificar_fila)

    def adicionar_imagem_grid(self, nome_arquivo, photo):
        """Cria os widgets da imagem e os posiciona no grid."""
        frame_item = ttk.Frame(self.scrollable_frame, borderwidth=2, relief="groove")
        frame_item.grid(row=self.grid_row, column=self.grid_col, padx=10, pady=10)

        # Imagem
        lbl_img = ttk.Label(frame_item, image=photo)
        lbl_img.pack()

        # Nome do arquivo
        lbl_texto = ttk.Label(frame_item, text=nome_arquivo, font=("Arial", 8), wraplength=200)
        lbl_texto.pack(pady=5)

        # Atualiza contadores do grid
        self.grid_col += 1
        if self.grid_col >= self.max_cols:
            self.grid_col = 0
            self.grid_row += 1

    def finalizar_carregamento(self, mensagem_status):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.lbl_progresso.pack_forget()
        self.status_var.set(mensagem_status)

if __name__ == "__main__":
    root = tk.Tk()
    app = VisualizadorPecas(root)
    root.mainloop()