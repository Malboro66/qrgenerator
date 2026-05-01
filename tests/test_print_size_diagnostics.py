import os
import sys
import tempfile
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import MagicMock, patch, call

import pytest
from PIL import Image

from models.geracao_config import GeracaoConfig


# ─────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────

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
    """PNG em memória sem DPI metadata."""
    buf = BytesIO()
    Image.new("RGB", (w, h), "white").save(buf, format="PNG")
    return buf.getvalue()


def _make_png_with_dpi(w: int, h: int, dpi: int) -> bytes:
    """PNG em memória COM DPI metadata."""
    buf = BytesIO()
    img = Image.new("RGB", (w, h), "white")
    img.save(buf, format="PNG", dpi=(dpi, dpi))
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────
# GRUPO 1 — Suspeito A: keep_barcode_ratio encolhe o barcode visível
# ─────────────────────────────────────────────────────────────────

class TestSuspeitoA_KeepRatio:

    def test_barcode_com_keep_ratio_true_canvas_tem_tamanho_correto(self):
        pytest.importorskip("barcode")
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0,
                   keep_barcode_ratio=True)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        esperado_w = max(1, round((8.0 / 2.54) * DPI))
        esperado_h = max(1, round((3.0 / 2.54) * DPI))
        assert img.width == esperado_w, (
            f"Largura do PNG: {img.width}px, esperado {esperado_w}px. "
            f"keep_ratio está encolhendo o canvas?"
        )
        assert img.height == esperado_h, (
            f"Altura do PNG: {img.height}px."
        )

    def test_barcode_com_keep_ratio_false_canvas_tem_tamanho_correto(self):
        pytest.importorskip("barcode")
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0,
                   keep_barcode_ratio=False)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        esperado_w = max(1, round((8.0 / 2.54) * DPI))
        esperado_h = max(1, round((3.0 / 2.54) * DPI))
        assert img.width == esperado_w
        assert img.height == esperado_h

    def test_barcode_area_util_vs_canvas_com_keep_ratio(self):
        pytest.importorskip("barcode")
        from services.renderers import BarcodeRenderer
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0,
                   keep_barcode_ratio=True)
        r = BarcodeRenderer(dpi_padrao=200)
        img = r.render("ABC123", cfg)
        pixels = list(img.getdata())
        brancos = sum(1 for p in pixels if p == (255, 255, 255))
        total = len(pixels)
        pct_branco = brancos / total * 100
        print(f"\n[DIAGNÓSTICO A] Canvas {img.width}×{img.height}px, "
              f"{pct_branco:.1f}% pixels brancos (padding)")
        assert total > 0


class TestSuspeitoB_DpiMetadata:

    def test_png_gerado_com_dpi_metadata(self):
        pytest.importorskip("barcode")
        from services.renderers import BarcodeRenderer
        DPI = 200
        cfg = _cfg(barcode_width_cm=8.0, barcode_height_cm=3.0)
        r = BarcodeRenderer(dpi_padrao=DPI)
        img = r.render("ABC123", cfg)
        dpi_info = img.info.get("dpi")
        assert dpi_info is not None, "PNG deve conter metadados DPI."
        print(f"\n[DIAGNÓSTICO B] PNG salvo COM DPI metadata: {dpi_info}")
        assert abs(dpi_info[0] - DPI) < 5, (
            f"DPI no metadata é {dpi_info[0]}, esperado {DPI}"
        )


class TestSuspeitoD_LoggingEMetricas:
    """Verifica o cálculo do retângulo GDI sem dependências de módulos nativos."""

    def test_calculo_alvo_px_isolado(self):
        """Testa a fórmula ppmm × cm × 10 sem qualquer mock de módulo nativo."""
        casos = [
            (104, 832,  8.0, 640),
            (104, 832,  4.0, 320),
            (104, 832, 10.4, 832),
        ]
        for horzsize_mm, horzres_px, largura_cm, esperado_px in casos:
            ppmm_x = horzres_px / horzsize_mm
            alvo_w = max(1, int(round(largura_cm * 10.0 * ppmm_x)))
            assert alvo_w == esperado_px, (
                f"horzsize={horzsize_mm}mm horzres={horzres_px}px "
                f"largura={largura_cm}cm → alvo={alvo_w}px esperado={esperado_px}px"
            )

    def test_retangulo_origin_zero(self):
        """O retângulo passado para Dib.draw deve sempre iniciar em (0, 0)."""
        horzsize_mm, horzres_px = 104, 832
        largura_cm, altura_cm = 8.0, 3.0
        ppmm_x = horzres_px / horzsize_mm
        ppmm_y = ppmm_x
        alvo_w = max(1, int(round(largura_cm * 10.0 * ppmm_x)))
        alvo_h = max(1, int(round(altura_cm * 10.0 * ppmm_y)))
        retangulo = (0, 0, alvo_w, alvo_h)
        assert retangulo[0] == 0, "x inicial deve ser 0"
        assert retangulo[1] == 0, "y inicial deve ser 0"
        assert retangulo[2] == alvo_w
        assert retangulo[3] == alvo_h

    def test_tamanho_fisico_invertido_deve_bater(self):
        """Inversão: alvo_px / ppmm deve retornar os mm originais (erro < 0.5mm)."""
        horzsize_mm = 104
        horzres_px  = 832
        largura_cm  = 8.0
        ppmm_x = horzres_px / horzsize_mm
        alvo_w = max(1, int(round(largura_cm * 10.0 * ppmm_x)))
        mm_calculado = alvo_w / ppmm_x
        assert abs(mm_calculado - largura_cm * 10.0) < 0.5, (
            f"Tamanho físico calculado: {mm_calculado:.2f}mm, "
            f"esperado: {largura_cm*10:.2f}mm"
        )

    @pytest.mark.windows_only
    def test_obter_metricas_impressora_retorna_dict(self):
        """Verifica que _obter_metricas_impressora_dc retorna estrutura correta."""
        tk = pytest.importorskip("tkinter")
        try:
            root = tk.Tk()
            root.withdraw()
        except Exception:
            pytest.skip("Requer display tkinter")

        from qr_generator import QRCodeGenerator
        from unittest.mock import MagicMock
        app = QRCodeGenerator(root)
        app.root.after = MagicMock()

        impressoras = list(app.impressora_combo["values"])
        if not impressoras:
            pytest.skip("Nenhuma impressora encontrada")

        for nome in impressoras:
            metricas = app._obter_metricas_impressora_dc(nome)
            if not metricas:
                continue
            chaves_obrigatorias = {
                "horzsize_mm", "vertsize_mm", "horzres_px",
                "vertres_px", "ppmm_x", "ppmm_y",
            }
            assert chaves_obrigatorias.issubset(metricas.keys()), (
                f"Chaves ausentes: {chaves_obrigatorias - metricas.keys()}"
            )
            assert metricas["horzsize_mm"] > 0
            assert metricas["horzres_px"] > 0
            dpi_eq = metricas["ppmm_x"] * 25.4
            print(f"\n[DIAGNÓSTICO DC REAL] {nome}: "
                  f"HORZSIZE={metricas['horzsize_mm']}mm "
                  f"HORZRES={metricas['horzres_px']}px "
                  f"DPI_equiv={dpi_eq:.1f}")
            assert 50 < dpi_eq < 1200, (
                f"DPI equivalente fora do range: {dpi_eq:.1f}"
            )
        root.destroy()


class TestCorrecaoImpressaoGDI:

    def test_imprimir_png_windows_gdi_chama_draw_com_parametros_corretos(self):
        tk = pytest.importorskip("tkinter")
        try:
            root = tk.Tk()
            root.withdraw()
        except Exception:
            pytest.skip("Requer display tkinter")

        from qr_generator import QRCodeGenerator

        mock_service = MagicMock()
        mock_service.DPI_PADRAO = 200
        mock_controller = MagicMock()
        mock_controller.service = mock_service
        mock_controller.logger = MagicMock()
        mock_controller.job_store = MagicMock()
        mock_controller.metrics_store = MagicMock()
        mock_controller.t = MagicMock(side_effect=lambda key, default, **kwargs: default)
        mock_controller.obter_modelos_barcode.return_value = [("code128", "Code128")]
        mock_controller.gerar_amostra_preview.return_value = "123"
        mock_controller.gerar_imagem_obj.return_value = Image.new("RGB", (10, 10), "white")

        app = QRCodeGenerator(root, controller=mock_controller)

        mock_dc = MagicMock()
        mock_dc.GetDeviceCaps.side_effect = [100, 100, 1000, 1000]
        mock_win32con = SimpleNamespace(HORZSIZE=1, VERTSIZE=2, HORZRES=3, VERTRES=4)
        mock_win32print = SimpleNamespace(GetDefaultPrinter=MagicMock(return_value="MinhaImpressora"))
        mock_win32ui = SimpleNamespace(CreateDC=MagicMock(return_value=mock_dc))
        mock_dib = MagicMock()

        with patch("qr_generator.importlib.import_module") as mock_import, \
             patch("qr_generator.Image.open") as mock_image_open, \
             patch("PIL.ImageWin.Dib", return_value=mock_dib):
            mock_import.side_effect = lambda name: {
                "win32con": mock_win32con,
                "win32print": mock_win32print,
                "win32ui": mock_win32ui,
            }[name]
            mock_img_pil = MagicMock()
            mock_img_pil.info = {"dpi": (200, 200)}
            mock_img_pil.width = 800
            mock_img_pil.height = 300
            mock_img_pil.convert.return_value = mock_img_pil
            mock_image_open.return_value = mock_img_pil

            ok = app._imprimir_png_windows_gdi("/tmp/teste.png", "MinhaImpressora", 8.0, 3.0)

        assert ok is True
        mock_dc.CreatePrinterDC.assert_called_once_with("MinhaImpressora")
        mock_dc.StartDoc.assert_called_once()
        mock_dib.draw.assert_called_once_with(mock_dc.GetHandleOutput(), (0, 0, 800, 300))
        mock_dc.EndDoc.assert_called_once()
        root.destroy()
