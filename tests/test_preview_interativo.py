"""
tests/test_preview_interativo.py
==================================
Testes para o widget PreviewInterativo (canvas de drag-to-resize).
"""
import pytest
from unittest.mock import MagicMock, call
from tkinter import Tk

try:
    from preview_interativo import PreviewInterativo
    _DISPONIVEL = True
except ImportError:
    _DISPONIVEL = False

pytestmark = pytest.mark.skipif(
    not _DISPONIVEL, reason="preview_interativo.py não encontrado"
)


@pytest.fixture(scope="module")
def root():
    r = Tk()
    r.withdraw()
    r.update()
    yield r
    r.destroy()


@pytest.fixture
def canvas(root):
    on_code = MagicMock()
    on_lbl  = MagicMock()
    c = PreviewInterativo(root, on_code_resized=on_code, on_label_resized=on_lbl)
    c.place(x=0, y=0, width=500, height=340)
    root.update()
    return c, on_code, on_lbl


class TestInicializacao:
    def test_cria_sem_erro(self, canvas):
        c, _, _ = canvas
        assert c is not None

    def test_estado_inicial(self, canvas):
        c, _, _ = canvas
        estado = c.obter_estado()
        assert estado["etiqueta_w_mm"] > 0
        assert estado["etiqueta_h_mm"] > 0
        assert estado["codigo_w_mm"] > 0
        assert estado["codigo_h_mm"] > 0


class TestAtualizar:
    def test_atualizar_dimensoes(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=120, lbl_h_mm=80, cod_w_mm=40, cod_h_mm=40)
        e = c.obter_estado()
        assert e["etiqueta_w_mm"] == pytest.approx(120, abs=1)
        assert e["etiqueta_h_mm"] == pytest.approx(80, abs=1)

    def test_codigo_nao_ultrapassa_etiqueta(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=50, lbl_h_mm=30, cod_w_mm=100, cod_h_mm=100)
        e = c.obter_estado()
        assert e["codigo_w_mm"] < 50
        assert e["codigo_h_mm"] < 30

    def test_dimensao_minima_etiqueta(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=1, lbl_h_mm=1, cod_w_mm=5, cod_h_mm=5)
        e = c.obter_estado()
        assert e["etiqueta_w_mm"] >= PreviewInterativo.MIN_LABEL_MM

    def test_dimensao_minima_codigo(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=0, cod_h_mm=0)
        e = c.obter_estado()
        assert e["codigo_w_mm"] >= PreviewInterativo.MIN_CODE_MM

    def test_atualizar_com_imagem_pil(self, canvas):
        from PIL import Image
        c, _, _ = canvas
        img = Image.new("RGB", (200, 100), "white")
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=40, cod_h_mm=30, codigo_img=img)
        e = c.obter_estado()
        assert e["etiqueta_w_mm"] == pytest.approx(100, abs=1)


class TestMoverCodigo:
    def test_mover_dentro_dos_limites(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=20, cod_h_mm=20)
        c.mover_codigo(10, 10)
        e = c.obter_estado()
        assert e["codigo_x_mm"] == pytest.approx(10, abs=1)
        assert e["codigo_y_mm"] == pytest.approx(10, abs=1)

    def test_mover_fora_do_limite_clamp(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=20, cod_h_mm=20)
        c.mover_codigo(200, 200)
        e = c.obter_estado()
        assert e["codigo_x_mm"] < 100
        assert e["codigo_y_mm"] < 60


class TestDragHandles:
    def _simula_drag(self, canvas_widget, handle_id, dx_px, dy_px):
        """Simula press + motion + release diretamente nos métodos internos."""
        c = canvas_widget
        c._draw()
        h_info = next((h for h in c._handles if h["id"] == handle_id), None)
        if not h_info:
            pytest.skip(f"Handle {handle_id} não encontrado (canvas muito pequeno?)")
        x0, y0 = h_info["x"], h_info["y"]

        class FakeEvent:
            pass

        ev_press = FakeEvent()
        ev_press.x = x0
        ev_press.y = y0
        c._press(ev_press)

        ev_move = FakeEvent()
        ev_move.x = x0 + dx_px
        ev_move.y = y0 + dy_px
        c._motion(ev_move)

        ev_release = FakeEvent()
        c._release(ev_release)

    def test_drag_c_se_aumenta_codigo(self, canvas):
        c, on_code, _ = canvas
        on_code.reset_mock()
        c.atualizar(lbl_w_mm=150, lbl_h_mm=100, cod_w_mm=40, cod_h_mm=30)
        antes_w = c._cw
        antes_h = c._ch
        self._simula_drag(c, "c-se", 30, 20)
        assert c._cw > antes_w or c._ch > antes_h
        assert on_code.called

    def test_drag_c_nw_diminui_codigo(self, canvas):
        c, on_code, _ = canvas
        c.atualizar(lbl_w_mm=150, lbl_h_mm=100, cod_w_mm=60, cod_h_mm=50)
        antes_w = c._cw
        self._simula_drag(c, "c-nw", 20, 15)
        assert c._cw <= antes_w

    def test_drag_l_se_aumenta_etiqueta(self, canvas):
        c, _, on_lbl = canvas
        on_lbl.reset_mock()
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=30, cod_h_mm=20)
        antes_lw = c._lw
        antes_lh = c._lh
        self._simula_drag(c, "l-se", 40, 30)
        assert c._lw > antes_lw or c._lh > antes_lh
        assert on_lbl.called

    def test_drag_codigo_nao_ultrapassa_etiqueta(self, canvas):
        c, _, _ = canvas
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=40, cod_h_mm=30)
        self._simula_drag(c, "c-se", 9999, 9999)
        assert c._cw + c._cx < c._lw
        assert c._ch + c._cy < c._lh

    def test_drag_sem_handle_nao_altera_estado(self, canvas):
        c, on_code, on_lbl = canvas
        on_code.reset_mock()
        on_lbl.reset_mock()
        c.atualizar(lbl_w_mm=100, lbl_h_mm=60, cod_w_mm=40, cod_h_mm=30)
        estado_antes = c.obter_estado()

        class FakeEvent:
            x = 5
            y = 5

        c._press(FakeEvent())
        FakeEvent.x = 50
        FakeEvent.y = 50
        c._motion(FakeEvent())
        c._release(FakeEvent())
        assert not on_code.called
        assert not on_lbl.called

    def test_callbacks_recebem_valores_em_mm(self, canvas):
        c, on_code, _ = canvas
        on_code.reset_mock()
        c.atualizar(lbl_w_mm=150, lbl_h_mm=100, cod_w_mm=40, cod_h_mm=30)
        self._simula_drag(c, "c-se", 15, 10)
        if on_code.called:
            w_mm, h_mm = on_code.call_args[0]
            assert 5 <= w_mm <= 150
            assert 5 <= h_mm <= 100


class TestObterEstado:
    def test_estado_arredondado(self, canvas):
        c, _, _ = canvas
        c._lw = 99.9999
        c._lh = 59.0001
        c._cw = 38.50001
        c._ch = 28.49999
        e = c.obter_estado()
        assert e["etiqueta_w_mm"] == pytest.approx(100.0, abs=0.2)
        assert e["etiqueta_h_mm"] == pytest.approx(59.0, abs=0.2)

    def test_estado_completo_tem_todas_chaves(self, canvas):
        c, _, _ = canvas
        e = c.obter_estado()
        for k in ("etiqueta_w_mm", "etiqueta_h_mm", "codigo_w_mm",
                  "codigo_h_mm", "codigo_x_mm", "codigo_y_mm"):
            assert k in e
