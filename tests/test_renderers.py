from types import SimpleNamespace
from unittest.mock import patch

import pytest
from PIL import Image

from services.renderers import BarcodeRenderer, ImageResizer, QRCodeRenderer


def _cfg(**kwargs):
    base = dict(
        qr_width_cm=4.0,
        qr_height_cm=4.0,
        barcode_width_cm=8.0,
        barcode_height_cm=3.0,
        keep_qr_ratio=True,
        keep_barcode_ratio=True,
        foreground="black",
        background="white",
        barcode_model="code128",
    )
    base.update(kwargs)
    return SimpleNamespace(**base)


class TestImageResizer:
    def test_resize_sem_ratio(self):
        img = Image.new("RGB", (100, 50), "white")
        out = ImageResizer.resize_with_ratio(img, 200, 80, keep_ratio=False)
        assert out.size == (200, 80)

    def test_resize_com_ratio_encaixa_canvas(self):
        img = Image.new("RGB", (100, 100), "white")
        out = ImageResizer.resize_with_ratio(img, 200, 100, keep_ratio=True)
        assert out.size == (200, 100)

    def test_resize_minimo_1px(self):
        img = Image.new("RGB", (10, 10), "white")
        out = ImageResizer.resize_with_ratio(img, 0, 0, keep_ratio=False)
        assert out.size == (1, 1)

    def test_resize_proporcional_nao_estica(self):
        img = Image.new("RGB", (200, 100), "white")
        out = ImageResizer.resize_with_ratio(img, 100, 100, keep_ratio=True)
        assert out.size == (100, 100)


class TestQRCodeRenderer:
    def test_render_retorna_pil_image(self):
        r = QRCodeRenderer()
        img = r.render("https://exemplo.com", _cfg())
        assert isinstance(img, Image.Image)

    def test_render_tamanho_correto(self):
        r = QRCodeRenderer()
        img = r.render("abc", _cfg(qr_width_cm=2.54, qr_height_cm=2.54))
        assert img.width > 0 and img.height > 0

    def test_render_modo_rgb(self):
        r = QRCodeRenderer()
        assert r.render("abc", _cfg()).mode == "RGB"

    def test_render_dado_longo(self):
        r = QRCodeRenderer()
        img = r.render("x" * 300, _cfg())
        assert img.width > 0

    def test_render_dado_vazio_nao_lanca(self):
        r = QRCodeRenderer()
        img = r.render("", _cfg())
        assert img.width > 0

    def test_render_caracteres_especiais(self):
        r = QRCodeRenderer()
        img = r.render("ação & çãõ", _cfg())
        assert img.width > 0


@pytest.mark.parametrize(
    "dado,modelo,deve_passar",
    [
        ("123456789012", "ean13", True),
        ("12345678901", "ean13", False),
        ("ABCDEF", "ean13", False),
        ("1234567", "ean8", True),
        ("123456", "ean8", False),
        ("12345678901", "upca", True),
        ("1234567890", "upca", False),
        ("12345678901231", "dun14", True),
        ("1234567890123", "dun14", False),
        ("12345678", "interleaved2of5", True),
        ("1234567", "interleaved2of5", False),
        ("ABCDEF12", "interleaved2of5", False),
    ],
)
def test_barcode_renderer_validacao(dado, modelo, deve_passar):
    if deve_passar:
        BarcodeRenderer.validar_modelo(dado, modelo)
    else:
        with pytest.raises(ValueError):
            BarcodeRenderer.validar_modelo(dado, modelo)


def test_modelo_desconhecido_lanca_runtime():
    r = BarcodeRenderer()
    with pytest.raises(RuntimeError, match="não suportado"):
        r.render("123", _cfg(barcode_model="modelo_inexistente"))


class TestBarcodeRendererRender:
    @pytest.mark.parametrize(
        "modelo,dado",
        [
            ("code128", "Hello123"),
            ("code39", "HELLO123"),
            ("ean13", "123456789012"),
            ("ean8", "1234567"),
            ("upca", "12345678901"),
            ("dun14", "12345678901231"),
            ("interleaved2of5", "12345678"),
        ],
    )
    def test_render_parametrizado(self, modelo, dado):
        pytest.importorskip("barcode")
        r = BarcodeRenderer()
        img = r.render(dado, _cfg(barcode_model=modelo))
        assert isinstance(img, Image.Image)
        assert img.width > 0

    def test_dun14_tem_moldura(self):
        pytest.importorskip("barcode")
        r = BarcodeRenderer()
        img = r.render("12345678901231", _cfg(barcode_model="dun14"))
        assert img.width > 0

    def test_render_tamanho_respeitado(self):
        pytest.importorskip("barcode")
        r = BarcodeRenderer(dpi_padrao=200)
        cfg = _cfg(keep_barcode_ratio=False, barcode_width_cm=5.0, barcode_height_cm=2.0, barcode_model="code128")
        img = r.render("ABC123", cfg)
        esperado = (int(round((5 / 2.54) * 200)), int(round((2 / 2.54) * 200)))
        assert img.size == esperado

    def test_fallback_para_reportlab_quando_pybarcode_falha(self):
        r = BarcodeRenderer()
        cfg = _cfg(barcode_model="code128")
        with patch.object(r, "_render_pybarcode", side_effect=ImportError("sem pybarcode")):
            try:
                img = r.render("ABC123", cfg)
                assert isinstance(img, Image.Image)
            except RuntimeError as exc:
                assert str(exc)

    def test_ambos_backends_falham_lanca_runtime_claro(self):
        r = BarcodeRenderer()
        cfg = _cfg(barcode_model="code128")
        with patch.object(r, "_render_pybarcode", side_effect=ImportError("x")), patch.object(
            r, "_render_reportlab", side_effect=RuntimeError("y")
        ):
            with pytest.raises(RuntimeError, match="python-barcode"):
                r.render("ABC123", cfg)
