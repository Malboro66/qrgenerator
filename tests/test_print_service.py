import tempfile
import types
import unittest
import importlib
from unittest.mock import MagicMock, patch

from PIL import Image

from services.print_service import _imprimir_png_windows_gdi


class TestPrintService(unittest.TestCase):
    @patch("services.print_service.importlib.import_module", side_effect=ImportError)
    def test_imprimir_sem_pywin32_retorna_false(self, _mock_import):
        resultado = _imprimir_png_windows_gdi("x.png", "", 4.0, 4.0)
        self.assertFalse(resultado)

    def test_imprimir_chama_gdi(self):
        real_import_module = importlib.import_module
        with tempfile.NamedTemporaryFile(suffix=".png") as tmp:
            Image.new("RGB", (20, 20), "white").save(tmp.name)

            hdc = MagicMock()
            hdc.GetDeviceCaps.side_effect = lambda cap: {
                1: 200,
                2: 200,
                3: 1600,
                4: 1600,
            }[cap]
            win32con = types.SimpleNamespace(HORZSIZE=1, VERTSIZE=2, HORZRES=3, VERTRES=4)
            win32print = types.SimpleNamespace(GetDefaultPrinter=MagicMock(return_value="IMP"))
            win32ui = types.SimpleNamespace(CreateDC=MagicMock(return_value=hdc))
            imagewin = types.SimpleNamespace(Dib=MagicMock(return_value=MagicMock(draw=MagicMock())))

            def fake_import(name):
                mods = {
                    "win32con": win32con,
                    "win32print": win32print,
                    "win32ui": win32ui,
                }
                if name in mods:
                    return mods[name]
                return real_import_module(name)

            with patch("services.print_service.importlib.import_module", side_effect=fake_import), patch(
                "PIL.ImageWin", imagewin, create=True
            ):
                ok = _imprimir_png_windows_gdi(tmp.name, "", 4.0, 4.0, dpi=200, logger=MagicMock())

            self.assertTrue(ok)
            hdc.CreatePrinterDC.assert_called_once_with("IMP")
            hdc.StartDoc.assert_called_once()
            hdc.EndDoc.assert_called_once()
            hdc.DeleteDC.assert_called_once()


if __name__ == "__main__":
    unittest.main()
