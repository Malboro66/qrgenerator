
import unittest
import os
import sys
import csv
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


class _ImmediateThread:
    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target = target
        self._args = args
        self._kwargs = kwargs or {}

    def start(self):
        if self._target:
            self._target(*self._args, **self._kwargs)


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
        self.app.qr_width_cm.set(3.0)
        self.app.atualizar_preview()
        self.assertIsNotNone(self.app.preview_image_ref)

    def test_atualizar_controles_formato(self):
        """Testa o alternar entre modos Texto e Numérico na UI."""
        # Garante layout inicial aplicado antes das assertivas
        self.app.atualizar_controles_formato()
        self.root.update_idletasks()
        
        # Modo texto padrão
        self.assertEqual(self.app.modo.get(), "texto")
        # Verifica qual container está ativo no layout
        self.assertEqual(self.app.texto_controls.winfo_manager(), "grid")
        self.assertEqual(self.app.numerico_controls.winfo_manager(), "")
        
        # Muda para numérico
        self.app.modo.set("numérico")
        self.app.atualizar_controles_formato()
        self.root.update_idletasks() # Processa o pack/pack_forget
        
        self.assertEqual(self.app.texto_controls.winfo_manager(), "")
        self.assertEqual(self.app.numerico_controls.winfo_manager(), "grid")

    # --- Testes de Carregamento de Dados ---

    def _criar_csv_temporario(self, dados):
        """Helper para criar CSV temporário."""
        fd, path = tempfile.mkstemp(suffix=".csv")
        try:
            with os.fdopen(fd, 'w', newline='', encoding='utf-8') as tmp:
                campos = list(dados.keys())
                writer = csv.DictWriter(tmp, fieldnames=campos)
                writer.writeheader()
                linhas = max(len(v) for v in dados.values()) if dados else 0
                for i in range(linhas):
                    writer.writerow({k: dados[k][i] for k in campos})
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
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.selecionar_arquivo()
            self.app.verificar_fila()

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
        with patch("qr_generator.threading.Thread", _ImmediateThread):
            self.app.selecionar_arquivo()
        self.app.verificar_fila()

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
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.selecionar_arquivo()
            self.app.verificar_fila()

            # Configura formato
            self.app.formato_saida.set("pdf")
            
            # Prepara fila para capturar mensagem (simulando thread)
            while not self.app.fila.empty():
                self.app.fila.get_nowait()

            # Ação: Iniciar geração pelo contrato atual
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.gerar_a_partir_da_tabela()
            
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
