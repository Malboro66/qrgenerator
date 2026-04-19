"""
tests/test_codigo_service.py
=============================
Testes unitários para CodigoService (camada de negócio).
"""
import pytest
from unittest.mock import patch, MagicMock
from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService


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


@pytest.fixture
def svc():
    return CodigoService()


class TestNormalizarDado:
    def test_modo_texto_sem_alteracao(self):
        assert CodigoService.normalizar_dado("abc", _cfg(modo="texto")) == "abc"

    def test_modo_numerico_com_prefixo_sufixo(self):
        cfg = _cfg(modo="numerico", prefixo="P-", sufixo="-FIM")
        assert CodigoService.normalizar_dado("123", cfg) == "P-123-FIM"

    def test_modo_numerico_sem_afixos(self):
        cfg = _cfg(modo="numerico", prefixo="", sufixo="")
        assert CodigoService.normalizar_dado("007", cfg) == "007"

    def test_converte_para_string(self):
        assert CodigoService.normalizar_dado(42, _cfg()) == "42"  # type: ignore[arg-type]


class TestSanitizarNomeArquivo:
    @pytest.mark.parametrize("entrada,esperado", [
        ("normal",         "normal"),
        ("com espaço",     "com espaço"),
        ('a/b\\c:*?"<>|',  "a_b_c____"),
        ("...oculto",      "oculto"),
        ("",               "fallback"),
        ("  ",             "fallback"),
    ])
    def test_casos(self, entrada, esperado):
        resultado = CodigoService.sanitizar_nome_arquivo(entrada, "fallback")
        assert resultado == esperado

    def test_fallback_quando_resultado_vazio(self):
        assert CodigoService.sanitizar_nome_arquivo("...", "fb") == "fb"


class TestValidarParametrosGeracao:
    def test_lista_vazia_lanca_valor(self):
        with pytest.raises(ValueError, match="Nenhum código válido"):
            CodigoService.validar_parametros_geracao([], _cfg())

    def test_limite_lote_excedido(self):
        codigos = ["a"] * 10
        with pytest.raises(ValueError, match="Limite excedido"):
            CodigoService.validar_parametros_geracao(codigos, _cfg(max_codigos_por_lote=5))

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
            CodigoService.validar_parametros_geracao(
                ["a"], _cfg(tipo_codigo="barcode", barcode_width_cm=41)
            )

    def test_todos_invalidos_lanca(self):
        codigos = ["" , "   "]
        with pytest.raises(ValueError, match="Todos os dados foram rejeitados"):
            CodigoService.validar_parametros_geracao(codigos, _cfg())

    def test_retorna_validos_e_invalidos(self):
        codigos = ["ok", "", "tambem_ok", " "]
        validos, invalidos = CodigoService.validar_parametros_geracao(codigos, _cfg())
        assert len(validos) == 2
        assert invalidos == 2

    def test_barcode_com_caracteres_de_controle_invalido(self):
        codigos = ["abc\x01def"]
        validos, invalidos = CodigoService.validar_parametros_geracao(
            codigos, _cfg(tipo_codigo="barcode")
        )
        assert invalidos == 1
        assert validos == []

    def test_dado_muito_longo_invalido(self):
        codigos = ["x" * 513]
        validos, invalidos = CodigoService.validar_parametros_geracao(codigos, _cfg())
        assert invalidos == 1

    def test_ean13_invalido_conta_como_invalido(self):
        codigos = ["nao-e-ean13"]
        validos, invalidos = CodigoService.validar_parametros_geracao(
            codigos, _cfg(tipo_codigo="barcode", barcode_model="ean13")
        )
        assert invalidos == 1

    def test_ean13_valido_conta_como_valido(self):
        codigos = ["123456789012"]
        validos, invalidos = CodigoService.validar_parametros_geracao(
            codigos, _cfg(tipo_codigo="barcode", barcode_model="ean13")
        )
        assert len(validos) == 1
        assert invalidos == 0

    def test_nao_modifica_valor_original_no_retorno(self):
        """Validos deve conter os valores BRUTOS originais, não normalizados."""
        codigos = ["123"]
        cfg = _cfg(modo="numerico", prefixo="P-", sufixo="-S")
        validos, _ = CodigoService.validar_parametros_geracao(codigos, cfg)
        assert validos[0] == "123"


class TestObterModelosBarcode:
    def test_retorna_lista_nao_vazia(self):
        modelos = CodigoService.obter_modelos_barcode()
        assert len(modelos) > 0

    def test_cada_item_e_tupla_chave_rotulo(self):
        for item in CodigoService.obter_modelos_barcode():
            chave, rotulo = item
            assert isinstance(chave, str)
            assert isinstance(rotulo, str)

    def test_code128_presente(self):
        chaves = [c for c, _ in CodigoService.obter_modelos_barcode()]
        assert "code128" in chaves


class TestGerarImagemObj:
    def test_qrcode_retorna_imagem(self, svc):
        from PIL import Image
        img = svc.gerar_imagem_obj("https://exemplo.com", _cfg())
        assert isinstance(img, Image.Image)

    def test_barcode_code128_retorna_imagem(self, svc):
        pytest.importorskip("barcode")
        from PIL import Image
        img = svc.gerar_imagem_obj("ABC123", _cfg(tipo_codigo="barcode"))
        assert isinstance(img, Image.Image)
