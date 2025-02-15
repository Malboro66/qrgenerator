import unittest
import os
import pandas as pd
from tkinter import Tk
from io import BytesIO
from unittest.mock import patch, MagicMock
import zipfile

# Import the QRCodeGenerator class from the main script
from qr_generator import QRCodeGenerator

class TestQRCodeGenerator(unittest.TestCase):
    def setUp(self):
        """Set up the test environment."""
        self.root = Tk()
        self.app = QRCodeGenerator(self.root)

    def tearDown(self):
        """Clean up after each test."""
        self.root.destroy()

    @patch("tkinter.filedialog.askdirectory")
    @patch("os.makedirs")
    @patch("os.path.exists", return_value=False)
    def test_gerar_imagens_png(self, mock_exists, mock_makedirs, mock_askdirectory):
        """Test generating PNG images."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["123", "456", "789"]
        self.app.gerar_imagens(codigos, "png", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.png")
            self.assertTrue(os.path.exists(caminho_arquivo))

    @patch("tkinter.filedialog.askdirectory")
    @patch("os.makedirs")
    @patch("os.path.exists", return_value=False)
    def test_gerar_imagens_svg(self, mock_exists, mock_makedirs, mock_askdirectory):
        """Test generating SVG images."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["abc", "def", "ghi"]
        self.app.gerar_imagens(codigos, "svg", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.svg")
            self.assertTrue(os.path.exists(caminho_arquivo))

    @patch("zipfile.ZipFile")
    @patch("tkinter.filedialog.askdirectory")
    def test_gerar_zip(self, mock_askdirectory, mock_zipfile):
        """Test generating a ZIP file."""
        mock_askdirectory.return_value = "/mock/directory"
        mock_zip_instance = MagicMock()
        mock_zipfile.return_value.__enter__.return_value = mock_zip_instance
        codigos = ["123", "456", "789"]
        self.app.gerar_zip(codigos, "/mock/directory")
        mock_zipfile.assert_called_once_with("/mock/directory/qrcodes.zip", "w")
        self.assertEqual(mock_zip_instance.writestr.call_count, len(codigos))

    @patch("reportlab.pdfgen.canvas.Canvas")
    def test_gerar_pdf(self, mock_canvas):
        """Test generating a PDF file."""
        mock_instance = MagicMock()
        mock_canvas.return_value = mock_instance
        codigos = ["123", "456", "789"]
        caminho_pdf = "/mock/output.pdf"
        self.app.gerar_pdf(codigos, caminho_pdf)
        mock_instance.save.assert_called_once()

    @patch("tkinter.filedialog.askdirectory")
    @patch("os.makedirs")
    @patch("os.path.exists", return_value=False)
    def test_nomeacao_automatica(self, mock_exists, mock_makedirs, mock_askdirectory):
        """Test automatic file naming based on data."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["produto1", "produto2", "produto3"]
        self.app.gerar_imagens(codigos, "png", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.png")
            self.assertTrue(os.path.exists(caminho_arquivo))

    @patch("tkinter.filedialog.askdirectory")
    @patch("os.makedirs")
    @patch("os.path.exists", return_value=False)
    def test_exportacao_multiplos_formatos(self, mock_exists, mock_makedirs, mock_askdirectory):
        """Test exporting QR Codes in multiple formats."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["123", "456", "789"]
        # Test PNG export
        self.app.formato_exportacao.set("png")
        self.app.gerar_qr_codes(codigos, "png", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.png")
            self.assertTrue(os.path.exists(caminho_arquivo))
        # Test SVG export
        self.app.formato_exportacao.set("svg")
        self.app.gerar_qr_codes(codigos, "svg", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.svg")
            self.assertTrue(os.path.exists(caminho_arquivo))
        # Test ZIP export
        self.app.formato_exportacao.set("zip")
        self.app.gerar_qr_codes(codigos, "zip", "/mock/directory")
        caminho_zip = os.path.join("/mock/directory", "qrcodes.zip")
        self.assertTrue(os.path.exists(caminho_zip))

    @patch("tkinter.filedialog.askdirectory")
    @patch("os.makedirs")
    @patch("os.path.exists", return_value=False)
    def test_exportacao_em_lote(self, mock_exists, mock_makedirs, mock_askdirectory):
        """Test batch export of QR Codes."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["batch1", "batch2", "batch3"]
        self.app.formato_exportacao.set("png")
        self.app.gerar_qr_codes(codigos, "png", "/mock/directory")
        for codigo in codigos:
            caminho_arquivo = os.path.join("/mock/directory", f"{codigo}.png")
            self.assertTrue(os.path.exists(caminho_arquivo))

    @patch("tkinter.messagebox.showinfo")
    @patch("tkinter.filedialog.askdirectory")
    def test_mostrar_sucesso_exportacao(self, mock_askdirectory, mock_messagebox):
        """Test success message display after export."""
        mock_askdirectory.return_value = "/mock/directory"
        codigos = ["success1", "success2"]
        self.app.gerar_qr_codes(codigos, "png", "/mock/directory")
        mock_messagebox.assert_called_once_with(
            "Sucesso",
            "QR Codes gerados com sucesso!\n\nLocal: /mock/directory"
        )

if __name__ == "__main__":
    unittest.main()