import csv
import os
import queue
import tempfile
import time
import unittest
import zipfile
from tkinter import Tk
from unittest.mock import MagicMock, patch

from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

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
    @classmethod
    def setUpClass(cls):
        cls.root = Tk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def setUp(self):
        self.app = QRCodeGenerator(self.root)
        self.root.after = MagicMock()

    def tearDown(self):
        while not self.app.fila.empty():
            try:
                self.app.fila.get_nowait()
            except queue.Empty:
                break

    def test_gerar_zip_imagem_png(self):
        codigos = ["zip_a", "zip_b"]
        with tempfile.TemporaryDirectory() as tmpdir:
            caminho_zip = os.path.join(tmpdir, "saida.zip")
            self.app.gerar_zip(codigos, caminho_zip)
            self.assertTrue(os.path.exists(caminho_zip))

            with zipfile.ZipFile(caminho_zip, "r") as zf:
                arquivos = zf.namelist()
                self.assertEqual(len(arquivos), 2)
                self.assertIn("zip_a.png", arquivos)
                self.assertIn("zip_b.png", arquivos)
                with zf.open("zip_a.png") as file:
                    with Image.open(file) as img:
                        self.assertEqual(img.format, "PNG")

    @patch("reportlab.pdfgen.canvas.Canvas")
    def test_gerar_pdf_mock_canvas(self, mock_canvas_class):
        mock_instance = MagicMock()
        mock_canvas_class.return_value = mock_instance
        self.app.qr_width_cm.set(5.0)
        self.app.qr_height_cm.set(2.5)

        self.app.gerar_pdf(["pdf_test"], "/tmp/nao_existe.pdf")

        mock_canvas_class.assert_called_with("/tmp/nao_existe.pdf", pagesize=A4)
        mock_instance.save.assert_called_once()
        self.assertTrue(mock_instance.drawImage.called)
        _, kwargs = mock_instance.drawImage.call_args
        self.assertAlmostEqual(kwargs["width"], 5.0 * 10 * mm)
        self.assertAlmostEqual(kwargs["height"], 2.5 * 10 * mm)

    @patch("reportlab.pdfgen.canvas.Canvas")
    def test_gerar_pdf_mock_canvas_respeita_tamanho_barcode(self, mock_canvas_class):
        mock_instance = MagicMock()
        mock_canvas_class.return_value = mock_instance
        self.app.tipo_codigo.set("barcode")
        self.app.barcode_model.set("code128")
        self.app.barcode_width_cm.set(9.0)
        self.app.barcode_height_cm.set(3.5)
        img_mock = Image.new("RGB", (200, 80), "white")
        with patch.object(self.app, "_gerar_imagem_obj", return_value=img_mock):
            self.app.gerar_pdf(["123456"], "/tmp/barcode.pdf")

        _, kwargs = mock_instance.drawImage.call_args
        self.assertAlmostEqual(kwargs["width"], 9.0 * 10 * mm)
        self.assertAlmostEqual(kwargs["height"], 3.5 * 10 * mm)

    def test_mensagens_de_progresso_na_fila(self):
        codigos = ["prog1", "prog2", "prog3"]
        with tempfile.TemporaryDirectory() as tmpdir:
            while not self.app.fila.empty():
                self.app.fila.get_nowait()

            self.app.gerar_imagens(codigos, "png", tmpdir)
            mensagens = []
            while True:
                try:
                    mensagens.append(self.app.fila.get_nowait())
                except queue.Empty:
                    break

            progressos = [m for m in mensagens if m["tipo"] == "progresso"]
            sucessos = [m for m in mensagens if m["tipo"] == "sucesso"]
            self.assertEqual(len(progressos), 3)
            self.assertEqual(len(sucessos), 1)
            self.assertEqual(sucessos[0]["caminho"], tmpdir)

    def test_atualizar_controles_formato(self):
        self.app.atualizar_controles_formato()
        self.root.update_idletasks()
        self.assertEqual(self.app.modo.get(), "texto")
        self.assertEqual(self.app.texto_controls.winfo_manager(), "grid")
        self.assertEqual(self.app.numerico_controls.winfo_manager(), "")

        self.app.modo.set("numerico")
        self.app.atualizar_controles_formato()
        self.root.update_idletasks()
        self.assertEqual(self.app.texto_controls.winfo_manager(), "")
        self.assertEqual(self.app.numerico_controls.winfo_manager(), "grid")

    def _criar_csv_temporario(self, dados):
        fd, path = tempfile.mkstemp(suffix=".csv")
        try:
            with os.fdopen(fd, "w", newline="", encoding="utf-8") as tmp:
                campos = list(dados.keys())
                writer = csv.DictWriter(tmp, fieldnames=campos)
                writer.writeheader()
                linhas = max(len(v) for v in dados.values()) if dados else 0
                for i in range(linhas):
                    writer.writerow({k: dados[k][i] for k in campos})
            return path
        except Exception:
            os.close(fd)
            raise

    @patch("tkinter.filedialog.askopenfilename")
    def test_selecionar_arquivo_csv_sucesso(self, mock_ask):
        dados = {"ID": [101, 102], "Valor": ["A", "B"]}
        path_csv = self._criar_csv_temporario(dados)
        try:
            mock_ask.return_value = path_csv
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.selecionar_arquivo()
            self.app.verificar_fila()

            self.assertEqual(self.app.arquivo_fonte, path_csv)
            self.assertEqual(self.app.column_combo.get(), "ID")
            self.assertEqual(str(self.app.generate_button["state"]), "normal")
        finally:
            os.unlink(path_csv)

    @patch("tkinter.messagebox.showerror")
    @patch("tkinter.filedialog.askopenfilename")
    def test_selecionar_arquivo_erro(self, mock_ask, mock_msg):
        mock_ask.return_value = "/caminho/que/nao/existe.xlsx"
        with patch("qr_generator.threading.Thread", _ImmediateThread):
            self.app.selecionar_arquivo()
        self.app.verificar_fila()

        self.assertTrue(mock_msg.called)
        self.assertEqual(str(self.app.column_combo["state"]), "disabled")
        self.assertEqual(str(self.app.generate_button["state"]), "disabled")

    @patch("tkinter.filedialog.asksaveasfilename")
    @patch("tkinter.filedialog.askdirectory")
    @patch("tkinter.filedialog.askopenfilename")
    def test_fluxo_completo_pdf(self, mock_open, _mock_dir, mock_save):
        dados = {"Codigos": ["X1", "X2"]}
        path_csv = self._criar_csv_temporario(dados)
        mock_open.return_value = path_csv

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
            path_pdf = tmp_pdf.name
        mock_save.return_value = path_pdf

        try:
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.selecionar_arquivo()
            self.app.verificar_fila()
            self.app.formato_saida.set("pdf")

            while not self.app.fila.empty():
                self.app.fila.get_nowait()

            with patch("qr_generator.threading.Thread", _ImmediateThread):
                self.app.gerar_a_partir_da_tabela()

            self.assertTrue(os.path.exists(path_pdf))
            msg = None
            start = time.time()
            while time.time() - start < 2:
                try:
                    m = self.app.fila.get(timeout=0.1)
                    if m["tipo"] == "sucesso":
                        msg = m
                        break
                except queue.Empty:
                    continue

            self.assertIsNotNone(msg, "Nenhuma mensagem de sucesso encontrada na fila")
            self.assertEqual(msg["tipo"], "sucesso")
        finally:
            if os.path.exists(path_csv):
                os.unlink(path_csv)
            if os.path.exists(path_pdf):
                os.unlink(path_pdf)


if __name__ == "__main__":
    unittest.main()