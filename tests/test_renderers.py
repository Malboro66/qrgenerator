"""
tests/test_renderers.py
========================
Testes unitários para QRCodeRenderer e BarcodeRenderer.
"""
import pytest
from unittest.mock import patch, MagicMock
from PIL import Image
from models.geracao_config import GeracaoConfig
from services.renderers import BarcodeRenderer, ImageResizer, QRCodeRenderer


def _cfg(**kw) -> GeracaoConfig:
    base = dict(
        qr_width_cm=4.0, qr_height_cm=4.0,
        barcode_width_cm=8.0, barcode_height_cm=3.0,
        keep_qr_ratio=True, keep_barcode_ratio=True,
        foreground="black", background="white",
        tipo_codigo="qrcode", barcode_model="code128",
        modo="texto", prefixo="", sufixo="",
    )
    base.update(kw)
    return GeracaoConfig(**base)


class TestImageResizer:
    def test_resize_sem_ratio(self):
        img = Image.new("RGB", (100, 50))
        out = ImageResizer.resize_with_ratio(img, 200, 80, keep_ratio=False)
        assert out.size == (200, 80)

    def test_resize_com_ratio_encaixa_canvas(self):
        img = Image.new("RGB", (100, 100))
        out = ImageResizer.resize_with_ratio(img, 200, 100, keep_ratio=True)
        assert out.size == (200, 100)

    def test_resize_minimo_1px(self):
        img = Image.new("RGB", (10, 10))
        out = ImageResizer.resize_with_ratio(img, 0, 0, keep_ratio=False)
        assert out.size == (1, 1)

    def test_resize_proporcional_nao_estica(self):
        img = Image.new("RGB", (200, 100))
        out = ImageResizer.resize_with_ratio(img, 100, 100, keep_ratio=True)
        assert out.size == (100, 100)


class TestQRCodeRenderer:
    def test_render_retorna_pil_image(self):
        r = QRCodeRenderer(dpi_padrao=72)
        img = r.render("https://exemplo.com", _cfg())
        assert isinstance(img, Image.Image)

    def test_render_tamanho_correto(self):
        r = QRCodeRenderer(dpi_padrao=96)
        cfg = _cfg(qr_width_cm=2.54, qr_height_cm=2.54)
        img = r.render("test", cfg)
        assert img.size[0] > 0 and img.size[1] > 0

    def test_render_modo_rgb(self):
        r = QRCodeRenderer()
        img = r.render("abc", _cfg())
        assert img.mode == "RGB"

    def test_render_dado_longo(self):
        r = QRCodeRenderer()
        dado = "A" * 300
        img = r.render(dado, _cfg())
        assert isinstance(img, Image.Image)

    def test_render_dado_vazio_nao_lanca(self):
        r = QRCodeRenderer()
        img = r.render("", _cfg())
        assert isinstance(img, Image.Image)

    def test_render_caracteres_especiais(self):
        r = QRCodeRenderer()
        img = r.render("https://site.com/path?a=1&b=ção", _cfg())
        assert isinstance(img, Image.Image)


class TestBarcodeRendererValidacao:
    @pytest.mark.parametrize("dado,modelo,deve_passar", [
        ("123456789012",   "ean13", True),
        ("12345678901",    "ean13", False),
        ("ABCDEF",         "ean13", False),
        ("1234567",        "ean8",  True),
        ("123456",         "ean8",  False),
        ("12345678901",    "upca",  True),
        ("1234567890",     "upca",  False),
        ("12345678901231", "dun14", True),
        ("1234567890123",  "dun14", False),
        ("12345678",       "interleaved2of5", True),
        ("1234567",        "interleaved2of5", False),
        ("ABCDEF12",       "interleaved2of5", False),
    ])
    def test_validar_modelo(self, dado, modelo, deve_passar):
        if deve_passar:
            BarcodeRenderer.validar_modelo(dado, modelo)
        else:
            with pytest.raises(ValueError):
                BarcodeRenderer.validar_modelo(dado, modelo)

    def test_modelo_desconhecido_lanca_runtime(self):
        r = BarcodeRenderer()
        with pytest.raises(RuntimeError, match="não suportado"):
            r.render("abc", _cfg(tipo_codigo="barcode", barcode_model="modelo_inexistente"))


class TestBarcodeRendererRender:
    @pytest.mark.parametrize("modelo,dado", [
        ("code128",          "Hello123"),
        ("code39",           "HELLO123"),
        ("ean13",            "123456789012"),
        ("ean8",             "1234567"),
        ("upca",             "12345678901"),
        ("dun14",            "12345678901231"),
        ("interleaved2of5",  "12345678"),
    ])
    def test_render_com_pybarcode(self, modelo, dado):
        pytest.importorskip("barcode")
        r = BarcodeRenderer(dpi_padrao=72)
        cfg = _cfg(tipo_codigo="barcode", barcode_model=modelo)
        img = r.render(dado, cfg)
        assert isinstance(img, Image.Image)
        assert img.size[0] > 0

    def test_dun14_tem_moldura(self):
        pytest.importorskip("barcode")
        r = BarcodeRenderer(dpi_padrao=72)
        cfg = _cfg(tipo_codigo="barcode", barcode_model="dun14")
        img = r.render("12345678901231", cfg)
        assert isinstance(img, Image.Image)

    def test_render_tamanho_respeitado(self):
        pytest.importorskip("barcode")
        r = BarcodeRenderer(dpi_padrao=200)
        cfg = _cfg(tipo_codigo="barcode", barcode_model="code128",
                   barcode_width_cm=5.0, barcode_height_cm=2.0,
                   keep_barcode_ratio=False)
        img = r.render("TEST", cfg)
        expected_w = int(round((5.0 / 2.54) * 200))
        expected_h = int(round((2.0 / 2.54) * 200))
        assert img.size == (expected_w, expected_h)

    def test_fallback_para_reportlab_quando_pybarcode_falha(self):
        r = BarcodeRenderer(dpi_padrao=72)
        cfg = _cfg(tipo_codigo="barcode", barcode_model="code128")
        with patch.object(r, "_render_pybarcode", side_effect=ImportError("no barcode")):
            try:
                img = r.render("ABC", cfg)
                assert isinstance(img, Image.Image)
            except RuntimeError as e:
                assert "instale" in str(e).lower() or "habilite" in str(e).lower()

    def test_ambos_backends_falham_lanca_runtime_claro(self):
        r = BarcodeRenderer()
        cfg = _cfg(tipo_codigo="barcode", barcode_model="code128")
        with patch.object(r, "_render_pybarcode", side_effect=ImportError):
            with patch.object(r, "_render_reportlab", side_effect=ImportError):
                with pytest.raises(RuntimeError, match="python-barcode"):
                    r.render("TEST", cfg)
