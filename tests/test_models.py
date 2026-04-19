import pytest

from models.geracao_config import GeracaoConfig


class TestGeracaoConfig:
    def _base(self, **kwargs):
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
        )
        data.update(kwargs)
        return GeracaoConfig(**data)

    def test_instancia_com_valores_padrao(self):
        cfg = self._base()
        assert cfg.max_codigos_por_lote == 5000
        assert cfg.max_tamanho_dado == 512

    def test_modo_numerico(self):
        cfg = self._base(modo="numerico", prefixo="P-", sufixo="-FIM")
        assert cfg.modo == "numerico"
        assert cfg.prefixo == "P-"
        assert cfg.sufixo == "-FIM"

    def test_campos_obrigatorios_presentes(self):
        cfg = self._base()
        obrigatorios = [
            "qr_width_cm",
            "qr_height_cm",
            "barcode_width_cm",
            "barcode_height_cm",
            "tipo_codigo",
            "barcode_model",
            "modo",
            "foreground",
            "background",
        ]
        for campo in obrigatorios:
            assert hasattr(cfg, campo), f"Campo obrigatório ausente: {campo}"

    def test_tipo_codigo_barcode(self):
        cfg = self._base(tipo_codigo="barcode", barcode_model="ean13")
        assert cfg.tipo_codigo == "barcode"
        assert cfg.barcode_model == "ean13"

    def test_limite_codigos_customizavel(self):
        cfg = self._base(max_codigos_por_lote=100)
        assert cfg.max_codigos_por_lote == 100

    @pytest.mark.parametrize(
        "modelo",
        ["code128", "ean13", "ean8", "upca", "code39", "code93", "dun14", "interleaved2of5", "codabar"],
    )
    def test_modelos_barcode_validos(self, modelo):
        cfg = self._base(tipo_codigo="barcode", barcode_model=modelo)
        assert cfg.barcode_model == modelo
