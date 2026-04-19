"""
tests/test_models.py
====================
Testes unitários para modelos de dados.
"""
import pytest
from dataclasses import FrozenInstanceError
from models.geracao_config import GeracaoConfig


class TestGeracaoConfig:
    def _config(self, **overrides) -> GeracaoConfig:
        defaults = dict(
            qr_width_cm=4.0, qr_height_cm=4.0,
            barcode_width_cm=8.0, barcode_height_cm=3.0,
            keep_qr_ratio=True, keep_barcode_ratio=True,
            foreground="black", background="white",
            tipo_codigo="qrcode", barcode_model="code128",
            modo="texto", prefixo="", sufixo="",
        )
        defaults.update(overrides)
        return GeracaoConfig(**defaults)

    def test_instancia_com_valores_padrao(self):
        c = self._config()
        assert c.max_codigos_por_lote == 5000
        assert c.max_tamanho_dado == 512

    def test_modo_numerico(self):
        c = self._config(modo="numerico", prefixo="P-", sufixo="-FIN")
        assert c.prefixo == "P-"
        assert c.sufixo == "-FIN"

    def test_campos_obrigatorios_presentes(self):
        c = self._config()
        for campo in (
            "qr_width_cm", "qr_height_cm", "barcode_width_cm",
            "barcode_height_cm", "tipo_codigo", "barcode_model",
            "modo", "foreground", "background",
        ):
            assert hasattr(c, campo), f"Campo obrigatório ausente: {campo}"

    def test_tipo_codigo_barcode(self):
        c = self._config(tipo_codigo="barcode", barcode_model="ean13")
        assert c.tipo_codigo == "barcode"
        assert c.barcode_model == "ean13"

    def test_limite_codigos_customizavel(self):
        c = self._config(max_codigos_por_lote=100)
        assert c.max_codigos_por_lote == 100

    @pytest.mark.parametrize("modelo", [
        "code128", "ean13", "ean8", "upca", "code39",
        "code93", "dun14", "interleaved2of5", "codabar",
    ])
    def test_modelos_barcode_validos(self, modelo):
        c = self._config(tipo_codigo="barcode", barcode_model=modelo)
        assert c.barcode_model == modelo
