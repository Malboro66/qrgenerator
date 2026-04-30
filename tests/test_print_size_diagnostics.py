import os
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest
from PIL import Image

from models.geracao_config import GeracaoConfig
from qr_generator import QRCodeGenerator


def _cfg(**kwargs):
    base = dict(
        qr_width_cm=4.0, qr_height_cm=4.0,
        barcode_width_cm=8.0, barcode_height_cm=3.0,
        keep_qr_ratio=True, keep_barcode_ratio=True,
        foreground="black", background="white",
        tipo_codigo="barcode", barcode_model="code128",
        modo="texto", prefixo="", sufixo="",
        etiqueta_width_mm=100.0, etiqueta_height_mm=60.0,
        max_codigos_por_lote=5000, max_tamanho_dado=512,
    )
    base.update(kwargs)
    return GeracaoConfig(**base)


def _make_png_bytes(w: int, h: int) -> bytes:
    buf = BytesIO()
    Image.new("RGB", (w, h), "white").save(buf, format="PNG")
    return buf.getvalue()


class TestSuspeitoA_KeepRatio:
    def test_barcode_com_keep_ratio_true_canvas_tem_tamanho_correto(self):
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0, keep_barcode_ratio=True)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        esperado_w = max(1, round((8.0 / 2.54) * DPI))
        esperado_h = max(1, round((3.0 / 2.54) * DPI))
        assert img.width == esperado_w
        assert img.height == esperado_h

    def test_barcode_com_keep_ratio_false_canvas_tem_tamanho_correto(self):
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0, keep_barcode_ratio=False)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        esperado_w = max(1, round((8.0 / 2.54) * DPI))
        esperado_h = max(1, round((3.0 / 2.54) * DPI))
        assert img.width == esperado_w
        assert img.height == esperado_h


class TestSuspeitoB_DpiMetadata:
    def test_png_gerado_sem_dpi_metadata(self):
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        img_lido = Image.open(buf)
        dpi_info = img_lido.info.get("dpi")
        if dpi_info is not None:
            assert abs(dpi_info[0] - DPI) < 5

    def test_tamanho_fisico_se_interpretado_como_96dpi(self):
        DPI_GERACAO = 200
        DPI_FALLBACK = 96
        largura_cm = 8.0
        altura_cm = 3.0
        px_w = round((largura_cm / 2.54) * DPI_FALLBACK)
        px_h = round((altura_cm / 2.54) * DPI_FALLBACK)
        cm_impresso_w = (px_w / DPI_FALLBACK) * 2.54
        cm_impresso_h = (px_h / DPI_FALLBACK) * 2.54
        assert cm_impresso_w < largura_cm
        assert cm_impresso_h < altura_cm


class TestSuspeitoC_CalculoGDI:
    @pytest.mark.parametrize("horzsize_mm,horzres_px,largura_cm,esperado_px", [
        (104, 832, 8.0, 640),
        (104, 832, 4.0, 320),
        (104, 832, 10.4, 832),
    ])
    def test_formula_cm_para_px(self, horzsize_mm, horzres_px, largura_cm, esperado_px):
        ppmm_x = horzres_px / horzsize_mm
        alvo_w = max(1, int(round(largura_cm * 10.0 * ppmm_x)))
        assert alvo_w == esperado_px


class TestSuspeitoD_LoggingEMetricas:
    def test_logging_diagnostico_antes_do_draw(self, tmp_path):
        img_path = tmp_path / "t.png"
        Image.new("RGB", (100, 50), "white").save(img_path)

        app = QRCodeGenerator.__new__(QRCodeGenerator)
        app.logger = MagicMock()
        app._formatar_excecao = lambda exc, m: f"{m}: {exc}"

        fake_dc = MagicMock()
        fake_dc.GetDeviceCaps.side_effect = [104, 60, 832, 480]

        fake_con = SimpleNamespace(HORZSIZE=1, VERTSIZE=2, HORZRES=3, VERTRES=4)
        fake_print = SimpleNamespace(GetDefaultPrinter=lambda: "ELGIN")
        fake_ui = SimpleNamespace(CreateDC=lambda: fake_dc)
        fake_dib = MagicMock()

        with patch("importlib.import_module", side_effect=[fake_con, fake_print, fake_ui]), \
             patch("PIL.ImageWin.Dib", return_value=fake_dib):
            ok = app._imprimir_png_windows_gdi(str(img_path), "ELGIN", 8.0, 3.0)

        assert ok is True
        app.logger.info.assert_called_once()
        assert fake_dib.draw.called

    def test_obter_metricas_dc_retorna_dict(self):
        app = QRCodeGenerator.__new__(QRCodeGenerator)
        hdc = MagicMock()
        hdc.GetDeviceCaps.side_effect = [104, 60, 832, 480, 203, 203, 0, 0, 832, 480]
        fake_con = SimpleNamespace(
            HORZSIZE=1, VERTSIZE=2, HORZRES=3, VERTRES=4,
            LOGPIXELSX=5, LOGPIXELSY=6,
            PHYSICALOFFSETX=7, PHYSICALOFFSETY=8,
            PHYSICALWIDTH=9, PHYSICALHEIGHT=10,
        )
        fake_ui = SimpleNamespace(CreateDC=lambda: hdc)
        with patch.dict("sys.modules", {"win32con": fake_con, "win32ui": fake_ui}):
            out = app._obter_metricas_impressora_dc("ELGIN")
        assert out["horzsize_mm"] == 104
        assert out["horzres_px"] == 832
        assert out["ppmm_x"] == 8.0

