from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from PIL import Image

from preview_interativo import PreviewInterativo


@pytest.fixture(scope="session")
def root():
    tk = pytest.importorskip("tkinter")
    try:
        r = tk.Tk()
    except Exception:
        pytest.skip("Requer display tkinter")
    r.withdraw()
    yield r
    r.destroy()


@pytest.fixture
@pytest.mark.ui
def canvas(root):
    on_code = Mock()
    on_label = Mock()
    c = PreviewInterativo(root, on_code_resized=on_code, on_label_resized=on_label, width=600, height=400)
    c.pack()
    c.update_idletasks()
    c._draw()
    yield c, on_code, on_label
    c.destroy()


class TestInicializacao:
    @pytest.mark.ui
    def test_cria_sem_erro(self, canvas):
        c, *_ = canvas
        assert c is not None

    @pytest.mark.ui
    def test_estado_inicial(self, canvas):
        c, *_ = canvas
        st = c.obter_estado()
        assert all(v > 0 for v in st.values())


class TestAtualizar:
    @pytest.mark.ui
    def test_atualizar_dimensoes(self, canvas):
        c, *_ = canvas
        c.atualizar(120, 80, 40, 40)
        st = c.obter_estado()
        assert st["etiqueta_w_mm"] == 120
        assert st["etiqueta_h_mm"] == 80

    @pytest.mark.ui
    def test_codigo_nao_ultrapassa_etiqueta(self, canvas):
        c, *_ = canvas
        c.atualizar(50, 30, 100, 100)
        st = c.obter_estado()
        assert st["codigo_w_mm"] < st["etiqueta_w_mm"]
        assert st["codigo_h_mm"] < st["etiqueta_h_mm"]

    @pytest.mark.ui
    def test_dimensao_minima_etiqueta(self, canvas):
        c, *_ = canvas
        c.atualizar(1, 1, 10, 10)
        assert c._lw == c.MIN_LABEL_MM and c._lh == c.MIN_LABEL_MM

    @pytest.mark.ui
    def test_dimensao_minima_codigo(self, canvas):
        c, *_ = canvas
        c.atualizar(100, 60, 0, 0)
        assert c._cw >= c.MIN_CODE_MM and c._ch >= c.MIN_CODE_MM

    @pytest.mark.ui
    def test_atualizar_com_imagem_pil(self, canvas):
        c, *_ = canvas
        c.atualizar(100, 60, 40, 20, codigo_img=Image.new("RGB", (200, 100), "black"))
        assert c._code_img is not None


class TestMoverCodigo:
    @pytest.mark.ui
    def test_mover_dentro_dos_limites(self, canvas):
        c, *_ = canvas
        c.atualizar(100, 60, 20, 20)
        c.mover_codigo(10, 10)
        assert c._cx == 10 and c._cy == 10

    @pytest.mark.ui
    def test_mover_fora_do_limite_clamp(self, canvas):
        c, *_ = canvas
        c.atualizar(100, 60, 20, 20)
        c.mover_codigo(200, 200)
        assert c._cx <= c._lw - c._cw
        assert c._cy <= c._lh - c._ch


def _simula_drag(c, handle_id, dx_px, dy_px):
    h = next(h for h in c._handles if h["id"] == handle_id)
    c._press(SimpleNamespace(x=h["x"], y=h["y"]))
    c._motion(SimpleNamespace(x=h["x"] + dx_px, y=h["y"] + dy_px))
    c._release(SimpleNamespace(x=h["x"] + dx_px, y=h["y"] + dy_px))


class TestDragHandles:
    @pytest.mark.ui
    def test_drag_c_se_aumenta_codigo(self, canvas):
        c, on_code, _ = canvas
        c.atualizar(100, 60, 20, 20)
        old = (c._cw, c._ch)
        _simula_drag(c, "c-se", 20, 20)
        assert c._cw > old[0] or c._ch > old[1]
        assert on_code.called

    @pytest.mark.ui
    def test_drag_c_nw_diminui_codigo(self, canvas):
        c, _, _ = canvas
        c.atualizar(100, 60, 30, 30)
        old = c._cw
        _simula_drag(c, "c-nw", 15, 15)
        assert c._cw <= old

    @pytest.mark.ui
    def test_drag_l_se_aumenta_etiqueta(self, canvas):
        c, _, on_label = canvas
        c.atualizar(100, 60, 20, 20)
        old = (c._lw, c._lh)
        _simula_drag(c, "l-se", 20, 20)
        assert c._lw > old[0] or c._lh > old[1]
        assert on_label.called

    @pytest.mark.ui
    def test_drag_codigo_nao_ultrapassa_etiqueta(self, canvas):
        c, *_ = canvas
        c.atualizar(100, 60, 20, 20)
        _simula_drag(c, "c-se", 500, 500)
        assert c._cw + c._cx < c._lw

    @pytest.mark.ui
    def test_drag_sem_handle_nao_altera_estado(self, canvas):
        c, on_code, on_label = canvas
        old = c.obter_estado().copy()
        c._press(SimpleNamespace(x=5, y=5))
        c._motion(SimpleNamespace(x=100, y=100))
        c._release(SimpleNamespace(x=100, y=100))
        assert c.obter_estado() == old
        assert not on_code.called and not on_label.called

    @pytest.mark.ui
    def test_callbacks_recebem_valores_em_mm(self, canvas):
        c, on_code, _ = canvas
        c.atualizar(100, 60, 20, 20)
        _simula_drag(c, "c-se", 20, 10)
        args, _ = on_code.call_args
        assert 5 <= args[0] <= 150
        assert 5 <= args[1] <= 150


class TestObterEstado:
    @pytest.mark.ui
    def test_estado_arredondado(self, canvas):
        c, *_ = canvas
        c._lw = 100.123
        st = c.obter_estado()
        assert st["etiqueta_w_mm"] == 100.1

    @pytest.mark.ui
    def test_estado_completo_tem_todas_chaves(self, canvas):
        c, *_ = canvas
        st = c.obter_estado()
        keys = {"etiqueta_w_mm", "etiqueta_h_mm", "codigo_w_mm", "codigo_h_mm", "codigo_x_mm", "codigo_y_mm"}
        assert set(st.keys()) == keys
