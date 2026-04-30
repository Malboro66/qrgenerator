import csv
import io
import os
import queue
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest
from PIL import Image

from qr_generator import EstadoAplicacao, QRCodeGenerator


class _ImmediateThread:
    """Stub que executa target() sincronamente."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self.target = target
        self.args = args
        self.kwargs = kwargs or {}

    def start(self):
        if self.target:
            self.target(*self.args, **self.kwargs)


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
def app(root):
    a = QRCodeGenerator(root)
    a.root.after = lambda *_args, **_kwargs: None
    while True:
        try:
            a.fila.get_nowait()
        except queue.Empty:
            break
    yield a
    while True:
        try:
            a.fila.get_nowait()
        except queue.Empty:
            break


def _csv_temp(dados):
    fd, caminho = tempfile.mkstemp(suffix=".csv")
    os.close(fd)
    cols = list(dados.keys())
    rows = zip(*[dados[c] for c in cols])
    with open(caminho, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for row in rows:
            w.writerow(row)
    return caminho


def _carregar_csv(app, dados):
    caminho = _csv_temp(dados)
    app._executar_carregamento(caminho)
    app.verificar_fila()
    return caminho


@pytest.mark.ui
class TestCarregamentoArquivo:
    def test_csv_carregado_popula_coluna(self, app):
        caminho = _carregar_csv(app, {"ID": ["1", "2"]})
        try:
            assert app.df is not None
            assert "ID" in app.column_combo["values"]
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_primeira_coluna_selecionada_automaticamente(self, app):
        caminho = _carregar_csv(app, {"SKU": ["A", "B"], "DESC": ["x", "y"]})
        try:
            assert app.column_combo.get() == "SKU"
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_botao_gerar_habilitado_apos_carga(self, app):
        caminho = _carregar_csv(app, {"ID": ["1"]})
        try:
            assert str(app.generate_button["state"]) == "normal"
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_arquivo_inexistente_mostra_erro(self, app):
        app._executar_carregamento("/tmp/nao-existe-xyz.csv")
        with patch("tkinter.messagebox.showerror") as m:
            app.verificar_fila()
            assert m.called

    def test_dialogo_cancelado_nao_altera_estado(self, app):
        app.df = None
        with patch("qr_generator.filedialog.askopenfilename", return_value=""):
            app.selecionar_arquivo()
        assert app.df is None


@pytest.mark.ui
class TestGeracaoPDF:
    def test_pdf_criado(self, app, tmp_path):
        pdf = tmp_path / "out.pdf"
        app.gerar_pdf(["abc", "def"], str(pdf), emitir_sucesso=False)
        assert pdf.exists() and pdf.stat().st_size > 1000

    def test_pdf_com_320_codigos(self, app, tmp_path):
        pdf = tmp_path / "many.pdf"
        cods = [f"id-{i}" for i in range(320)]
        app.gerar_pdf(cods, str(pdf), emitir_sucesso=False)
        assert pdf.exists() and pdf.stat().st_size > 5000

    def test_pdf_dimensoes_qr_corretas(self, app, tmp_path):
        app.qr_width_cm.set("5.0")
        pdf = tmp_path / "size.pdf"
        app.gerar_pdf(["abc"], str(pdf), emitir_sucesso=False)
        assert pdf.exists()


@pytest.mark.ui
class TestGeracaoPNG:
    def test_png_individuais_criados(self, app, tmp_path):
        app.gerar_imagens(["1", "2", "3"], "png", str(tmp_path), emitir_sucesso=False)
        pngs = sorted(tmp_path.glob("*.png"))
        assert len(pngs) == 3, f"Esperado 3 PNGs, encontrado: {pngs}"

    def test_png_e_imagem_valida(self, app, tmp_path):
        app.gerar_imagens(["1"], "png", str(tmp_path), emitir_sucesso=False)
        p = next(tmp_path.glob("*.png"))
        with Image.open(p) as im:
            assert im.format == "PNG"

    def test_nome_arquivo_sanitizado(self, app, tmp_path):
        app.gerar_imagens(["a/b", "c\\d"], "png", str(tmp_path), emitir_sucesso=False)
        nomes = [p.name for p in tmp_path.glob("*.png")]
        assert all("/" not in n and "\\" not in n for n in nomes)

    def test_nomes_duplicados_com_sufixo_numerico(self, app, tmp_path):
        app.gerar_imagens(["dup", "dup", "dup"], "png", str(tmp_path), emitir_sucesso=False)
        nomes = [p.name for p in tmp_path.glob("*.png")]
        assert len(nomes) == 3 and len(set(nomes)) == 3


@pytest.mark.ui
class TestGeracaoZIP:
    def test_zip_criado_com_pngs(self, app, tmp_path):
        z = tmp_path / "codes.zip"
        app.gerar_zip(["1", "2", "3"], str(z))
        assert z.exists()
        with zipfile.ZipFile(z) as zf:
            nomes = zf.namelist()
        assert len(nomes) == 3 and all(n.endswith(".png") for n in nomes)

    def test_zip_conteudo_valido(self, app, tmp_path):
        z = tmp_path / "codes.zip"
        app.gerar_zip(["1"], str(z))
        with zipfile.ZipFile(z) as zf:
            nome = zf.namelist()[0]
            data = zf.read(nome)
        with Image.open(io.BytesIO(data)) as im:
            assert im.format == "PNG"


@pytest.mark.ui
class TestGeracaoSVG:
    def test_svg_criado(self, app, tmp_path):
        app.tipo_codigo.set("qrcode")
        app.gerar_imagens(["x"], "svg", str(tmp_path), emitir_sucesso=False)
        assert len(list(tmp_path.glob("*.svg"))) == 1

    def test_svg_com_barcode_lanca_erro(self, app, tmp_path):
        pytest.importorskip("barcode")
        app.tipo_codigo.set("barcode")
        app.barcode_model.set("code128")
        app.barcode_disponivel = True
        with pytest.raises(RuntimeError, match="SVG"):
            app.gerar_imagens(["x"], "svg", str(tmp_path), emitir_sucesso=False)


@pytest.mark.ui
class TestCancelamento:
    def test_estado_volta_para_ready_apos_cancelamento(self, app):
        app.df = [{"a": 1}]
        app.fila.put({"tipo": "cancelado", "msg": "ok"})
        with patch("tkinter.messagebox.showinfo"):
            app.verificar_fila()
        assert app.estado_atual in {EstadoAplicacao.READY, EstadoAplicacao.IDLE}


@pytest.mark.ui
class TestFilaDeMensagens:
    def test_mensagem_progresso_atualiza_barra(self, app):
        app.fila.put({"tipo": "progresso", "atual": 5, "total": 10, "codigo": "x"})
        app.verificar_fila()
        assert int(app.progress_bar["value"]) == 5

    def test_mensagem_erro_mostra_dialogo(self, app):
        app.fila.put({"tipo": "erro", "msg": "falha"})
        with patch("tkinter.messagebox.showerror") as m:
            app.verificar_fila()
            assert m.called

    def test_mensagem_sucesso_mostra_info(self, app):
        app._total_planejado = 1
        app.fila.put({"tipo": "sucesso", "caminho": "/tmp/arquivo"})
        with patch("tkinter.messagebox.showinfo"):
            app.verificar_fila()
        assert "/tmp/arquivo" in app.resumo_caminho_var.get()

    def test_multiplas_mensagens_processadas_em_sequencia(self, app):
        app.fila.put({"tipo": "progresso", "atual": 3, "total": 10, "codigo": "x"})
        app.fila.put({"tipo": "progresso", "atual": 5, "total": 10, "codigo": "y"})
        app.verificar_fila()
        assert int(app.progress_bar["value"]) == 5


@pytest.mark.ui
class TestModoNumerico:
    def test_prefixo_sufixo_aplicados_na_geracao(self, app, tmp_path):
        app.modo.set("numerico")
        app.prefixo_numerico.set("P-")
        app.sufixo_numerico.set("-S")
        capt = {}

        def fake(dado, cfg=None):
            capt["dado"] = dado
            return Image.new("RGB", (30, 30), "white")

        with patch.object(app, "_gerar_imagem_obj", side_effect=fake):
            app.gerar_imagens(["123"], "png", str(tmp_path), emitir_sucesso=False)
        assert capt["dado"] == "P-123-S"

    def test_modo_texto_nao_aplica_afixos(self, app, tmp_path):
        app.modo.set("texto")
        app.prefixo_numerico.set("P-")
        app.sufixo_numerico.set("-S")
        capt = {}

        def fake(dado, cfg=None):
            capt["dado"] = dado
            return Image.new("RGB", (30, 30), "white")

        with patch.object(app, "_gerar_imagem_obj", side_effect=fake):
            app.gerar_imagens(["abc"], "png", str(tmp_path), emitir_sucesso=False)
        assert capt["dado"] == "abc"


@pytest.mark.ui
class TestStepper:
    def test_etapa2_bloqueada_sem_arquivo(self, app):
        app.df = None
        with patch("tkinter.messagebox.showwarning") as m:
            app._definir_etapa(2)
            assert m.called

    def test_etapa3_bloqueada_sem_coluna(self, app):
        app.df = [{"x": 1}]
        app.column_combo.set("")
        with patch("tkinter.messagebox.showwarning") as m:
            app._definir_etapa(3)
            assert m.called

    def test_etapa1_sempre_acessivel(self, app):
        app._definir_etapa(1)
        assert app.etapa_atual == 1


@pytest.mark.ui
class TestBuildConfig:
    def test_config_basica_criada(self, app):
        cfg = app._build_config()
        assert cfg.qr_width_cm > 0

    def test_barcode_indisponivel_lanca_valor(self, app):
        app.tipo_codigo.set("barcode")
        app.barcode_disponivel = False
        with pytest.raises(ValueError, match="indisponível"):
            app._build_config()

    def test_parse_float_virgula_como_separador(self, app):
        assert app._parse_float_input("3,5", "x") == 3.5

    def test_parse_float_invalido_lanca_valor(self, app):
        with pytest.raises(ValueError):
            app._parse_float_input("abc", "x")


@pytest.mark.ui
class TestImpressao:
    def test_imprimir_nao_windows_lanca_runtime(self, app):
        with patch("sys.platform", "linux"):
            with pytest.raises(RuntimeError, match="Windows"):
                app.imprimir_codigos(["1"])

    def test_limpar_temporarios_remove_pastas_velhas(self, app, tmp_path):
        d = tmp_path / "velha"
        d.mkdir()
        app._arquivos_temporarios_impressao = [(str(d), time.time() - 10000)]
        app._limpar_arquivos_temporarios_impressao(idade_min_segundos=100)
        assert not d.exists()

    def test_limpar_temporarios_mantem_pastas_novas(self, app, tmp_path):
        d = tmp_path / "nova"
        d.mkdir()
        app._arquivos_temporarios_impressao = [(str(d), time.time())]
        app._limpar_arquivos_temporarios_impressao(idade_min_segundos=100)
        assert d.exists()


@pytest.mark.ui
class TestFormatoIncompativel:
    def test_svg_barcode_formata_para_png(self, app):
        app.tipo_codigo.set("barcode")
        app.formato_saida.set("svg")
        app._ajustar_formato_incompativel()
        assert app.formato_saida.get() == "png"

    def test_svg_qrcode_mantido(self, app):
        app.tipo_codigo.set("qrcode")
        app.formato_saida.set("svg")
        app._ajustar_formato_incompativel()
        assert app.formato_saida.get() == "svg"
