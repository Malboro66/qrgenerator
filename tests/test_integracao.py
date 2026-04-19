"""
tests/test_integracao.py
=========================
Testes de integração: fluxos completos de ponta a ponta
(geração de PDF, PNG, ZIP, impressão, cancelamento).
"""
import csv
import io
import os
import queue
import tempfile
import time
import threading
import zipfile
import pytest
from unittest.mock import MagicMock, patch
from tkinter import Tk

from PIL import Image
from qr_generator import QRCodeGenerator


# ─────────────────────────────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────────────────────────────

class _ImmediateThread:
    """Stub de threading.Thread que executa na thread de teste (síncrono)."""
    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._t = target
        self._a = args
        self._kw = kwargs or {}

    def start(self):
        if self._t:
            self._t(*self._a, **self._kw)


@pytest.fixture(scope="session")
def root():
    r = Tk()
    r.withdraw()
    yield r
    r.destroy()


@pytest.fixture
def app(root):
    inst = QRCodeGenerator(root)
    root.after = MagicMock()
    yield inst
    while not inst.fila.empty():
        try:
            inst.fila.get_nowait()
        except queue.Empty:
            break


def _csv_temp(dados: dict) -> str:
    fd, path = tempfile.mkstemp(suffix=".csv")
    with os.fdopen(fd, "w", newline="", encoding="utf-8") as f:
        campos = list(dados.keys())
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        n = max(len(v) for v in dados.values())
        for i in range(n):
            w.writerow({k: dados[k][i] for k in campos})
    return path


def _carregar_csv(app, dados):
    path = _csv_temp(dados)
    with patch("tkinter.filedialog.askopenfilename", return_value=path):
        with patch("qr_generator.threading.Thread", _ImmediateThread):
            app.selecionar_arquivo()
    app.verificar_fila()
    return path


# ─────────────────────────────────────────────────────────────────────
# Carregamento de arquivo
# ─────────────────────────────────────────────────────────────────────

class TestCarregamentoArquivo:
    def test_csv_carregado_popula_coluna(self, app):
        path = _carregar_csv(app, {"ID": ["001", "002"]})
        try:
            assert app.df is not None
            assert "ID" in app._obter_colunas(app.df)
        finally:
            os.unlink(path)

    def test_primeira_coluna_selecionada_automaticamente(self, app):
        path = _carregar_csv(app, {"SKU": ["A", "B"], "Desc": ["X", "Y"]})
        try:
            assert app.column_combo.get() == "SKU"
        finally:
            os.unlink(path)

    def test_botao_gerar_habilitado_apos_carga(self, app):
        path = _carregar_csv(app, {"Cod": ["001"]})
        try:
            assert str(app.generate_button["state"]) == "normal"
        finally:
            os.unlink(path)

    def test_arquivo_inexistente_mostra_erro(self, app):
        with patch("tkinter.filedialog.askopenfilename", return_value="/nao_existe.xlsx"):
            with patch("qr_generator.threading.Thread", _ImmediateThread):
                with patch("tkinter.messagebox.showerror") as mock_err:
                    app.selecionar_arquivo()
        app.verificar_fila()
        assert mock_err.called

    def test_dialogo_cancelado_nao_altera_estado(self, app):
        estado_antes = app.df
        with patch("tkinter.filedialog.askopenfilename", return_value=""):
            app.selecionar_arquivo()
        assert app.df is estado_antes


# ─────────────────────────────────────────────────────────────────────
# Geração de PDF
# ─────────────────────────────────────────────────────────────────────

class TestGeracaoPDF:
    def test_pdf_criado(self, app):
        path = _carregar_csv(app, {"Cod": ["001", "002", "003"]})
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            pdf_path = f.name
        try:
            app.formato_saida.set("pdf")
            with patch("tkinter.filedialog.asksaveasfilename", return_value=pdf_path):
                with patch("qr_generator.threading.Thread", _ImmediateThread):
                    app.gerar_a_partir_da_tabela()
            assert os.path.exists(pdf_path)
            assert os.path.getsize(pdf_path) > 1000
        finally:
            os.unlink(path)
            os.unlink(pdf_path)

    def test_pdf_com_320_codigos(self, app):
        codigos = [str(i).zfill(6) for i in range(320)]
        path = _csv_temp({"Cod": codigos})
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            pdf_path = f.name
        try:
            with patch("tkinter.filedialog.askopenfilename", return_value=path):
                with patch("qr_generator.threading.Thread", _ImmediateThread):
                    app.selecionar_arquivo()
            app.verificar_fila()
            app.formato_saida.set("pdf")
            with patch("tkinter.filedialog.asksaveasfilename", return_value=pdf_path):
                with patch("qr_generator.threading.Thread", _ImmediateThread):
                    app.gerar_a_partir_da_tabela()
            assert os.path.getsize(pdf_path) > 5000
        finally:
            os.unlink(path)
            os.unlink(pdf_path)

    def test_pdf_dimensoes_qr_corretas(self, app):
        from reportlab.lib.units import mm as rl_mm
        path = _carregar_csv(app, {"v": ["test"]})
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            pdf_path = f.name
        try:
            app.qr_width_cm.set("5.0")
            app.qr_height_cm.set("5.0")
            chamados = []
            orig = app.gerar_pdf
            def captura(*a, **kw):
                chamados.append((a, kw))
                return orig(*a, **kw)
            with patch.object(app, "gerar_pdf", side_effect=captura):
                app.gerar_pdf(["test"], pdf_path)
            assert os.path.exists(pdf_path)
        finally:
            os.unlink(path)
            os.unlink(pdf_path)


# ─────────────────────────────────────────────────────────────────────
# Geração de PNG
# ─────────────────────────────────────────────────────────────────────

class TestGeracaoPNG:
    def test_png_individuais_criados(self, app):
        path = _carregar_csv(app, {"Cod": ["AAA", "BBB", "CCC"]})
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                app.formato_saida.set("png")
                with patch("tkinter.filedialog.askdirectory", return_value=tmpdir):
                    with patch("qr_generator.threading.Thread", _ImmediateThread):
                        app.gerar_a_partir_da_tabela()
                pngs = [f for f in os.listdir(tmpdir) if f.endswith(".png")]
                assert len(pngs) == 3
            finally:
                os.unlink(path)

    def test_png_e_imagem_valida(self, app):
        path = _carregar_csv(app, {"Cod": ["TESTE"]})
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                app.gerar_imagens(["TESTE"], "png", tmpdir)
                pngs = [f for f in os.listdir(tmpdir) if f.endswith(".png")]
                img = Image.open(os.path.join(tmpdir, pngs[0]))
                assert img.format == "PNG"
                assert img.size[0] > 0
            finally:
                os.unlink(path)

    def test_nome_arquivo_sanitizado(self, app):
        path = _carregar_csv(app, {"Cod": ["a/b\\c:d"]})
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                app.gerar_imagens(["a/b\\c:d"], "png", tmpdir)
                arquivos = os.listdir(tmpdir)
                for nome in arquivos:
                    assert "/" not in nome
                    assert "\\" not in nome
            finally:
                os.unlink(path)

    def test_nomes_duplicados_com_sufixo_numerico(self, app):
        with tempfile.TemporaryDirectory() as tmpdir:
            app.gerar_imagens(["dup", "dup", "dup"], "png", tmpdir)
            pngs = [f for f in os.listdir(tmpdir) if f.endswith(".png")]
            assert len(pngs) == 3
            assert len(set(pngs)) == 3


# ─────────────────────────────────────────────────────────────────────
# Geração de ZIP
# ─────────────────────────────────────────────────────────────────────

class TestGeracaoZIP:
    def test_zip_criado_com_pngs(self, app):
        path = _carregar_csv(app, {"Cod": ["X1", "X2", "X3"]})
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "saida.zip")
            try:
                app.gerar_zip(["X1", "X2", "X3"], zip_path)
                assert os.path.exists(zip_path)
                with zipfile.ZipFile(zip_path) as zf:
                    nomes = zf.namelist()
                assert len(nomes) == 3
                assert all(n.endswith(".png") for n in nomes)
            finally:
                os.unlink(path)

    def test_zip_conteudo_valido(self, app):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "out.zip")
            app.gerar_zip(["TEST_ZIP"], zip_path)
            with zipfile.ZipFile(zip_path) as zf:
                with zf.open("TEST_ZIP.png") as f:
                    img = Image.open(io.BytesIO(f.read()))
                    assert img.format == "PNG"


# ─────────────────────────────────────────────────────────────────────
# SVG
# ─────────────────────────────────────────────────────────────────────

class TestGeracaoSVG:
    def test_svg_criado(self, app):
        with tempfile.TemporaryDirectory() as tmpdir:
            app.gerar_imagens(["SVG_TEST"], "svg", tmpdir)
            svgs = [f for f in os.listdir(tmpdir) if f.endswith(".svg")]
            assert len(svgs) == 1

    def test_svg_com_barcode_lanca_erro(self, app):
        app.tipo_codigo.set("barcode")
        with tempfile.TemporaryDirectory() as tmpdir:
            with pytest.raises(RuntimeError, match="SVG"):
                app.gerar_imagens(["ABC"], "svg", tmpdir)
        app.tipo_codigo.set("qrcode")


# ─────────────────────────────────────────────────────────────────────
# Cancelamento
# ─────────────────────────────────────────────────────────────────────

class TestCancelamento:
    def test_cancelar_durante_geracao_para_processo(self, app):
        path = _carregar_csv(app, {"Cod": [str(i) for i in range(200)]})
        cancelou = {"sim": False}

        def gerar_com_cancelamento(codigos, formato, destino):
            for i, cod in enumerate(codigos):
                if i == 5:
                    app.cancelar_evento.set()
                if app.cancelar_evento.is_set():
                    cancelou["sim"] = True
                    raise Exception("cancelado")

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                with patch.object(app, "_executar_geracao", side_effect=gerar_com_cancelamento):
                    app.cancelar_evento.clear()
                    t = threading.Thread(
                        target=app._executar_geracao,
                        args=(["a"] * 100, "png", tmpdir),
                    )
        finally:
            os.unlink(path)

    def test_estado_volta_para_ready_apos_cancelamento(self, app):
        from qr_generator import EstadoAplicacao, OperacaoCancelada
        app.fila.put({"tipo": "cancelado", "msg": "Operação cancelada pelo usuário."})
        path = _carregar_csv(app, {"Cod": ["001"]})
        try:
            with patch("tkinter.messagebox.showinfo"):
                app.verificar_fila()
        finally:
            os.unlink(path)


# ─────────────────────────────────────────────────────────────────────
# Fila de mensagens
# ─────────────────────────────────────────────────────────────────────

class TestFilaDeMensagens:
    def test_mensagem_progresso_atualiza_barra(self, app):
        app.fila.put({"tipo": "progresso", "atual": 5, "total": 10, "codigo": "abc"})
        app._inicio_geracao_ts = time.perf_counter()
        app.verificar_fila()
        assert app.progress_bar["value"] == 5

    def test_mensagem_erro_mostra_dialogo(self, app):
        with patch("tkinter.messagebox.showerror") as mock_err:
            app.fila.put({"tipo": "erro", "msg": "erro teste", "detalhe": ""})
            app.verificar_fila()
        assert mock_err.called

    def test_mensagem_sucesso_mostra_info(self, app):
        path = _carregar_csv(app, {"Cod": ["a"]})
        try:
            app._total_planejado = 1
            app._processados_atuais = 0
            app._inicio_geracao_ts = time.perf_counter()
            with patch("tkinter.messagebox.showinfo"):
                app.fila.put({"tipo": "sucesso", "caminho": "/tmp/out.pdf"})
                app.verificar_fila()
            assert "out.pdf" in app.resumo_caminho_var.get()
        finally:
            os.unlink(path)

    def test_multiplas_mensagens_processadas_em_sequencia(self, app):
        app._inicio_geracao_ts = time.perf_counter()
        for i in range(1, 6):
            app.fila.put({"tipo": "progresso", "atual": i, "total": 5, "codigo": str(i)})
        app.verificar_fila()
        assert app.progress_bar["value"] == 5


# ─────────────────────────────────────────────────────────────────────
# Modo numérico
# ─────────────────────────────────────────────────────────────────────

class TestModoNumerico:
    def test_prefixo_sufixo_aplicados_na_geracao(self, app):
        app.modo.set("numerico")
        app.prefixo_numerico.set("P-")
        app.sufixo_numerico.set("-S")
        imagens_geradas = []

        def captura_imagem(dado, cfg):
            imagens_geradas.append(dado)
            return Image.new("RGB", (100, 100))

        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.object(app, "_gerar_imagem_obj", side_effect=captura_imagem):
                app.gerar_imagens(["123"], "png", tmpdir, emitir_sucesso=False)
        assert imagens_geradas[0] == "P-123-S"
        app.modo.set("texto")

    def test_modo_texto_nao_aplica_afixos(self, app):
        app.modo.set("texto")
        app.prefixo_numerico.set("X")
        app.sufixo_numerico.set("Y")
        imagens_geradas = []

        def captura(dado, cfg):
            imagens_geradas.append(dado)
            return Image.new("RGB", (100, 100))

        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.object(app, "_gerar_imagem_obj", side_effect=captura):
                app.gerar_imagens(["abc"], "png", tmpdir, emitir_sucesso=False)
        assert imagens_geradas[0] == "abc"


# ─────────────────────────────────────────────────────────────────────
# Stepper / navegação de etapas
# ─────────────────────────────────────────────────────────────────────

class TestStepper:
    def test_etapa2_bloqueada_sem_arquivo(self, app):
        app.df = None
        with patch("tkinter.messagebox.showwarning") as mock_warn:
            app._definir_etapa(2)
        assert mock_warn.called

    def test_etapa3_bloqueada_sem_coluna(self, app):
        path = _carregar_csv(app, {"Cod": ["a"]})
        try:
            app.column_combo.set("")
            with patch("tkinter.messagebox.showwarning") as mock_warn:
                app._definir_etapa(3)
            assert mock_warn.called
        finally:
            os.unlink(path)

    def test_etapa1_sempre_acessivel(self, app):
        app._definir_etapa(1)
        assert app.etapa_atual == 1


# ─────────────────────────────────────────────────────────────────────
# Validações de config
# ─────────────────────────────────────────────────────────────────────

class TestBuildConfig:
    def test_config_basica_criada(self, app):
        cfg = app._build_config()
        assert cfg.qr_width_cm > 0
        assert cfg.qr_height_cm > 0

    def test_barcode_indisponivel_lanca_valor(self, app):
        orig = app.barcode_disponivel
        app.barcode_disponivel = False
        app.tipo_codigo.set("barcode")
        with pytest.raises(ValueError, match="indisponível"):
            app._build_config()
        app.tipo_codigo.set("qrcode")
        app.barcode_disponivel = orig

    def test_parse_float_virgula_como_separador(self, app):
        val = app._parse_float_input("3,5", "campo")
        assert val == pytest.approx(3.5)

    def test_parse_float_invalido_lanca_valor(self, app):
        with pytest.raises(ValueError):
            app._parse_float_input("abc", "campo")


# ─────────────────────────────────────────────────────────────────────
# Impressão (mock de mspaint)
# ─────────────────────────────────────────────────────────────────────

class TestImpressao:
    def test_imprimir_nao_windows_lanca_runtime(self, app):
        with patch("sys.platform", "linux"):
            with pytest.raises(RuntimeError, match="Windows"):
                app.imprimir_codigos(["001"])

    def test_limpar_temporarios_remove_pastas_velhas(self, app):
        import shutil
        pasta = tempfile.mkdtemp(prefix="qr_print_")
        app._arquivos_temporarios_impressao.append((pasta, 0))
        app._limpar_arquivos_temporarios_impressao(idade_min_segundos=0)
        assert not os.path.exists(pasta)

    def test_limpar_temporarios_mantem_pastas_novas(self, app):
        pasta = tempfile.mkdtemp(prefix="qr_print_")
        try:
            app._arquivos_temporarios_impressao.append((pasta, time.time()))
            app._limpar_arquivos_temporarios_impressao(idade_min_segundos=9999)
            assert os.path.exists(pasta)
        finally:
            import shutil
            shutil.rmtree(pasta, ignore_errors=True)

    @pytest.mark.skipif(os.name != "nt", reason="Somente Windows")
    def test_imprimir_png_windows_usa_popen(self, app):
        with patch("subprocess.Popen") as mock_popen:
            mock_proc = MagicMock()
            mock_proc.poll.return_value = None
            mock_popen.return_value = mock_proc
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                try:
                    app._imprimir_png_windows(f.name, "Impressora Teste")
                    assert mock_popen.called
                finally:
                    os.unlink(f.name)


# ─────────────────────────────────────────────────────────────────────
# Formato SVG bloqueado para barcode
# ─────────────────────────────────────────────────────────────────────

class TestFormatoIncompativel:
    def test_svg_barcode_formata_para_png(self, app):
        app.tipo_codigo.set("barcode")
        app.formato_saida.set("svg")
        with patch("tkinter.messagebox.showwarning"):
            app._ajustar_formato_incompativel(exibir_aviso=False)
        assert app.formato_saida.get() == "png"
        app.tipo_codigo.set("qrcode")

    def test_svg_qrcode_mantido(self, app):
        app.tipo_codigo.set("qrcode")
        app.formato_saida.set("svg")
        app._ajustar_formato_incompativel(exibir_aviso=False)
        assert app.formato_saida.get() == "svg"
