from types import SimpleNamespace

import pytest
from PIL import Image

from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService


def _cfg(**kwargs):
    data = dict(
        qr_width_cm=4.0,
        qr_height_cm=4.0,
        barcode_width_cm=8.0,
        barcode_height_cm=3.0,
        keep_qr_ratio=True,
        keep_barcode_ratio=True,
        foreground="black",
        background="white",
        tipo_codigo="qrcode",
        barcode_model="code128",
        modo="texto",
        prefixo="",
        sufixo="",
        max_codigos_por_lote=5000,
        max_tamanho_dado=512,
    )
    data.update(kwargs)
    return GeracaoConfig(**data)


class TestNormalizarDado:
    def test_modo_texto_sem_alteracao(self):
        assert CodigoService.normalizar_dado("abc", _cfg(modo="texto")) == "abc"

    def test_modo_numerico_com_prefixo_sufixo(self):
        assert CodigoService.normalizar_dado("123", _cfg(modo="numerico", prefixo="P-", sufixo="-FIM")) == "P-123-FIM"

    def test_modo_numerico_sem_afixos(self):
        assert CodigoService.normalizar_dado("007", _cfg(modo="numerico")) == "007"

    def test_converte_para_string(self):
        assert CodigoService.normalizar_dado(42, _cfg()) == "42"


@pytest.mark.parametrize(
    "entrada,esperado",
    [
        ("normal", "normal"),
        ("com espaço", "com espaço"),
        ('a/b\\c:*?"<>|', "a_b_c_______"),
        ("...oculto", "oculto"),
        ("", "fallback"),
        ("  ", "fallback"),
    ],
)
def test_sanitizar_nome_arquivo_parametrizado(entrada, esperado):
    assert CodigoService.sanitizar_nome_arquivo(entrada, "fallback") == esperado


def test_fallback_quando_resultado_vazio():
    assert CodigoService.sanitizar_nome_arquivo("...", "fb") == "fb"


class TestValidarParametrosGeracao:
    def test_lista_vazia_lanca_valor(self):
        with pytest.raises(ValueError, match="Nenhum código válido"):
            CodigoService.validar_parametros_geracao([], _cfg())

    def test_limite_lote_excedido(self):
        with pytest.raises(ValueError, match="Limite excedido"):
            CodigoService.validar_parametros_geracao(["a"] * 10, _cfg(max_codigos_por_lote=5))

    def test_tamanho_qr_invalido_zero(self):
        with pytest.raises(ValueError, match="Tamanho de QR inválido"):
            CodigoService.validar_parametros_geracao(["a"], _cfg(qr_width_cm=0))

    def test_tamanho_qr_invalido_negativo(self):
        with pytest.raises(ValueError):
            CodigoService.validar_parametros_geracao(["a"], _cfg(qr_height_cm=-1))

    def test_tamanho_qr_acima_do_limite(self):
        with pytest.raises(ValueError, match="30 cm"):
            CodigoService.validar_parametros_geracao(["a"], _cfg(qr_width_cm=31))

    def test_tamanho_barcode_acima_do_limite(self):
        with pytest.raises(ValueError, match="40x20"):
            CodigoService.validar_parametros_geracao(["a"], _cfg(barcode_width_cm=41))

    def test_todos_invalidos_lanca(self):
        with pytest.raises(ValueError, match="Todos os dados foram rejeitados"):
            CodigoService.validar_parametros_geracao(["", "   "], _cfg())

    def test_retorna_validos_e_invalidos(self):
        validos, invalidos = CodigoService.validar_parametros_geracao(["ok", "", "tambem_ok", " "], _cfg())
        assert len(validos) == 2
        assert invalidos == 2

    def test_barcode_com_caracteres_de_controle_invalido(self):
        with pytest.raises(ValueError, match="Todos os dados"):
            CodigoService.validar_parametros_geracao(["abc\x01def"], _cfg(tipo_codigo="barcode"))

    def test_dado_muito_longo_invalido(self):
        with pytest.raises(ValueError, match="Todos os dados"):
            CodigoService.validar_parametros_geracao(["x" * 513], _cfg())

    def test_ean13_invalido_conta_como_invalido(self):
        with pytest.raises(ValueError, match="Todos os dados"):
            CodigoService.validar_parametros_geracao(
                ["nao-e-ean13"], _cfg(tipo_codigo="barcode", barcode_model="ean13")
            )

    def test_ean13_valido_conta_como_valido(self):
        validos, invalidos = CodigoService.validar_parametros_geracao(
            ["123456789012"], _cfg(tipo_codigo="barcode", barcode_model="ean13")
        )
        assert len(validos) == 1
        assert invalidos == 0

    def test_nao_modifica_valor_original_no_retorno(self):
        validos, _ = CodigoService.validar_parametros_geracao(
            ["123"], _cfg(tipo_codigo="barcode", modo="numerico", prefixo="P-", barcode_model="code128")
        )
        assert validos[0] == "123"


class TestObterModelosBarcode:
    def test_retorna_lista_nao_vazia(self):
        out = CodigoService.obter_modelos_barcode()
        assert out

    def test_cada_item_e_tupla_chave_rotulo(self):
        out = CodigoService.obter_modelos_barcode()
        assert all(isinstance(i, tuple) and len(i) == 2 for i in out)

    def test_code128_presente(self):
        out = dict(CodigoService.obter_modelos_barcode())
        assert "code128" in out


class TestGerarImagemObj:
    def test_qrcode_retorna_imagem(self):
        s = CodigoService()
        img = s.gerar_imagem_obj("https://exemplo.com", _cfg(tipo_codigo="qrcode"))
        assert isinstance(img, Image.Image)

    def test_barcode_code128_retorna_imagem(self):
        pytest.importorskip("barcode")
        s = CodigoService()
        img = s.gerar_imagem_obj("ABC123", _cfg(tipo_codigo="barcode", barcode_model="code128"))
        assert isinstance(img, Image.Image)
