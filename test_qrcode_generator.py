import unittest
import os
import sys
import pandas as pd
import tempfile
import zipfile
import time
import queue
from unittest.mock import patch, MagicMock, call
from PIL import Image
from tkinter import Tk

# Importação necessária para corrigir o erro "A4 is not defined"
from reportlab.lib.pagesizes import A4

from qr_generator import QRCodeGenerator

class TestQRCodeGenerator(unittest.TestCase):
    """Suite de testes robusta para o Gerador de QR Codes."""

    @classmethod
    def setUpClass(cls):
        """Configurações executadas uma vez antes de todos os testes."""
        # Evita que o Tkinter mostre janelas reais ou bloqueie a execução
        cls.root = Tk()
        cls.root.withdraw()  # Esconde a janela principal

    def setUp(self):
        """Configurações executadas antes de cada teste."""
        # Recria a aplicação para garantir estado limpo
        self.app = QRCodeGenerator(self.root)
        
        # Mock do root.after para evitar loops de eventos infinitos durante testes
        # Isso permite chamar verificar_fila sem travar o teste
        self.root.after = MagicMock()

    def tearDown(self):
        """Limpeza após cada teste."""
        # Limpa a fila para não vazar mensagens entre testes
        while not self.app.fila.empty():
            try:
                self.app.fila.get_nowait()
            except:
                pass

    @classmethod
    def tearDownClass(cls):
        """Limpeza final após todos os testes."""
        cls.root.destroy()

    # --- Testes de Geração de Imagens (I/O Real) ---

    def test_gerar_imagens_png_sucesso(self):
        """Testa a geração de arquivos PNG e verifica a integridade da imagem."""
        codigos = ["teste1", "teste2", "teste3"]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            # Executa a geração
            self.app.gerar_imagens(codigos, "png", tmpdir)
            
            # Verifica se os arquivos foram criados
            for codigo in codigos:
                caminho = os.path.join(tmpdir, f"{codigo}.png")
                self.assertTrue(os.path.exists(caminho), f"Arquivo {caminho} não foi criado.")
                
                # Verifica se é um PNG válido usando PIL
                with Image.open(caminho) as img:
                    self.assertEqual(img.format, "PNG")
                    # QR codes podem ser '1' (bw), 'L' (grayscale) ou 'RGB' dependendo das cores e versões do Pillow
                    self.assertIn(img.mode, ["1", "RGB", "L"])

    def test_gerar_imagens_svg_sucesso(self):
        """Testa a geração de arquivos SVG."""
        codigos = ["svg_01"]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            self.app.gerar_imagens(codigos, "svg", tmpdir)
            
            caminho = os.path.join(tmpdir, f"{codigos[0]}.svg")
            self.assertTrue(os.path.exists(caminho))
            
            # Verifica conteúdo do arquivo
            with open(caminho, 'r', encoding='utf-8') as f:
                content = f.read()
                self.assertIn('<svg', content)
                self.assertIn('xmlns', content)
                # Dados do QR não estão em texto no SVG, são formas geométricas. Não testamos a string do dado.

    def test_gerar_zip_conteudo(self):
        """Testa a geração de ZIP e verifica se os arquivos internos estão corretos."""
        codigos = ["zip_a", "zip_b"]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            caminho_zip = os.path.join(tmpdir, "qrcodes.zip")
            
            # Executa
            self.app.gerar_zip(codigos, caminho_zip)
            
            # Verifica criação do ZIP
            self.assertTrue(os.path.exists(caminho_zip))
            
            # Verifica conteúdo interno
            with zipfile.ZipFile(caminho_zip, 'r') as zf:
                arquivos = zf.namelist()
                self.assertEqual(len(arquivos), 2)
                self.assertIn("zip_a.png", arquivos)
                self.assertIn("zip_b.png", arquivos)
                
                # Verifica se o arquivo dentro do zip é uma imagem válida
                with zf.open("zip_a.png") as file:
                    with Image.open(file) as img:
                        self.assertEqual(img.format, "PNG")

    # --- Testes de Geração de PDF ---

    @patch('reportlab.pdfgen.canvas.Canvas')
    def test_gerar_pdf_mock_canvas(self, mock_canvas_class):
        """Testa o fluxo de geração de PDF usando mock para o Canvas (rápido e isolado)."""
        mock_instance = MagicMock()
        mock_canvas_class.return_value = mock_instance
        
        codigos = ["pdf_test"]
        caminho_pdf = "/tmp/nao_existe.pdf" # O path não importa pois mockamos o Canvas
        
        self.app.gerar_pdf(codigos, caminho_pdf)
        
        # Verifica se o canvas foi criado com o caminho certo e pagesize correto (A4)
        # Nota: O código usa 'from reportlab.lib.pagesizes import A4'
        mock_canvas_class.assert_called_with(caminho_pdf, pagesize=A4)
        
        # Verifica se save foi chamado
        mock_instance.save.assert_called_once()
        
        # Verifica se drawImage foi chamado (sinal de que tentou desenhar)
        self.assertTrue(mock_instance.drawImage.called)

    def test_gerar_pdf_real_file(self):
        """Teste de integração: Gera um PDF real e verifica se existe e tem tamanho."""
        codigos = ["real_pdf"]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            caminho = os.path.join(tmpdir, "saida.pdf")
            
            self.app.gerar_pdf(codigos, caminho)
            
            self.assertTrue(os.path.exists(caminho))
            self.assertGreater(os.path.getsize(caminho), 1000) # PDF deve ter algum tamanho

    # --- Testes de Fila e Concorrência ---

    def test_mensagens_de_progresso_na_fila(self):
        """Verifica se as mensagens de progresso são enviadas para a fila corretamente."""
        codigos = ["prog1", "prog2", "prog3"]
        
        with tempfile.TemporaryDirectory() as tmpdir:
            # Limpa a fila explicitamente antes do teste
            while not self.app.fila.empty():
                self.app.fila.get_nowait()
            
            self.app.gerar_imagens(codigos, "png", tmpdir)
            
            # Coleta mensagens
            mensagens = []
            # Pequeno timeout para garantir que a thread/processamento terminou
            # (Como aqui é síncrono direto, deve ser imediato, mas boa prática)
            try:
                while True:
                    msg = self.app.fila.get_nowait()
                    mensagens.append(msg)
            except:
                pass # Fila vazia
            
            # Deve ter 3 mensagens de progresso e 1 de sucesso
            progressos = [m for m in mensagens if m['tipo'] == 'progresso']
            sucessos = [m for m in mensagens if m['tipo'] == 'sucesso']
            
            self.assertEqual(len(progressos), 3)
            self.assertEqual(len(sucessos), 1)
            self.assertEqual(sucessos[0]['caminho'], tmpdir)

    # --- Testes de Interface e Preview ---

    def test_preview_geracao(self):
        """Testa se o preview gera uma imagem sem crashar."""
        # Deve gerar preview inicial
        self.app.atualizar_preview()
        self.assertIsNotNone(self.app.preview_image_ref)
        
        # Alterar cor e atualizar
        self.app.qr_foreground_color.set("blue")
        self.app.atualizar_preview()
        self.assertIsNotNone(self.app.preview_image_ref)
        
        # Alterar tamanho
        self.app.qr_size.set(300)
        self.app.atualizar_preview()
        self.assertIsNotNone(self.app.preview_image_ref)

    def test_atualizar_controles_formato(self):
        """Testa o alternar entre modos Texto e Numérico na UI."""
        # Atualiza a interface para garantir que o layout foi processado
        self.root.update_idletasks()
        
        # Modo texto padrão
        self.assertEqual(self.app.modo.get(), "texto")
        # Verifica se os controles de texto estão visíveis (winfo_ismapped)
        self.assertTrue(self.app.texto_controls.winfo_ismapped())
        self.assertFalse(self.app.numerico_controls.winfo_ismapped())
        
        # Muda para numérico
        self.app.modo.set("numerico")
        self.app.atualizar_controles_formato()
        self.root.update_idletasks() # Processa o pack/pack_forget
        
        self.assertFalse(self.app.texto_controls.winfo_ismapped())
        self.assertTrue(self.app.numerico_controls.winfo_ismapped())

    # --- Testes de Carregamento de Dados ---

    def _criar_csv_temporario(self, dados):
        """Helper para criar CSV temporário."""
        fd, path = tempfile.mkstemp(suffix=".csv")
        try:
            with os.fdopen(fd, 'w') as tmp:
                df = pd.DataFrame(dados)
                df.to_csv(tmp, index=False)
            return path
        except:
            os.close(fd)
            raise

    @patch('tkinter.filedialog.askopenfilename')
    def test_selecionar_arquivo_csv_sucesso(self, mock_ask):
        """Testa o carregamento de um CSV real."""
        dados = {'ID': [101, 102], 'Valor': ['A', 'B']}
        path_csv = self._criar_csv_temporario(dados)
        
        try:
            mock_ask.return_value = path_csv
            self.app.selecionar_arquivo()
            
            self.assertEqual(self.app.arquivo_fonte, path_csv)
            self.assertEqual(self.app.column_combo.get(), 'ID')
            # Converte para string para comparação segura entre versões Tk
            self.assertEqual(str(self.app.generate_button['state']), 'normal')
        finally:
            os.unlink(path_csv)

    @patch('tkinter.messagebox.showerror')
    @patch('tkinter.filedialog.askopenfilename')
    def test_selecionar_arquivo_erro(self, mock_ask, mock_msg):
        """Testa tratamento de erro ao carregar arquivo inválido."""
        mock_ask.return_value = "/caminho/que/nao/existe.xlsx"
        self.app.selecionar_arquivo()
        
        # Deve ter mostrado erro
        self.assertTrue(mock_msg.called)
        # Converte para string para comparação segura
        self.assertEqual(str(self.app.column_combo['state']), 'disabled')
        self.assertEqual(str(self.app.generate_button['state']), 'disabled')

    # --- Testes de Integração Fluxo Completo ---

    @patch('tkinter.filedialog.asksaveasfilename')
    @patch('tkinter.filedialog.askdirectory')
    @patch('tkinter.filedialog.askopenfilename')
    def test_fluxo_completo_pdf(self, mock_open, mock_dir, mock_save):
        """Testa o fluxo: Selecionar arquivo -> Gerar PDF -> Verificar Sucesso."""
        # 1. Setup CSV
        dados = {'Codigos': ['X1', 'X2']}
        path_csv = self._criar_csv_temporario(dados)
        mock_open.return_value = path_csv
        
        # 2. Setup Save PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
            path_pdf = tmp_pdf.name
        mock_save.return_value = path_pdf
        
        try:
            # Ação: Selecionar arquivo
            self.app.selecionar_arquivo()
            
            # Configura formato
            self.app.formato_exportacao.set("pdf")
            
            # Prepara fila para capturar mensagem (simulando thread)
            while not self.app.fila.empty():
                self.app.fila.get_nowait()

            # Ação: Iniciar geração
            # Vamos chamar processar_dados diretamente
            coluna = self.app.column_combo.get()
            
            # Mockamos a UI thread update para não chamar root.after real
            self.app.processar_dados(coluna)
            
            # Verifica se o arquivo foi criado
            self.assertTrue(os.path.exists(path_pdf))
            
            # Verifica mensagem de sucesso na fila (drenando a fila até achar sucesso)
            msg = None
            start = time.time()
            while time.time() - start < 2:
                try:
                    m = self.app.fila.get(timeout=0.1)
                    if m['tipo'] == 'sucesso':
                        msg = m
                        break
                except queue.Empty:
                    continue
            
            self.assertIsNotNone(msg, "Nenhuma mensagem de sucesso encontrada na fila")
            self.assertEqual(msg['tipo'], 'sucesso')
            
        finally:
            if os.path.exists(path_csv): os.unlink(path_csv)
            if os.path.exists(path_pdf): os.unlink(path_pdf)

if __name__ == "__main__":
    unittest.main()