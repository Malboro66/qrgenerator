from models.geracao_config import GeracaoConfig
from services.data_importer import DataImporter
from services.renderers import BarcodeRenderer, QRCodeRenderer


class CodigoService:
    """Camada de negócio orquestrando importação, validação e renderização."""

    DPI_PADRAO = 200
    BARCODE_MODELOS_SUPORTADOS = BarcodeRenderer.MODELOS_SUPORTADOS

    def __init__(self):
        self.data_importer = DataImporter()
        self.qr_renderer = QRCodeRenderer(self.DPI_PADRAO)
        self.barcode_renderer = BarcodeRenderer(self.DPI_PADRAO)

    @staticmethod
    def formatar_excecao(exc: Exception, contexto: str) -> str:
        return DataImporter.formatar_excecao(exc, contexto)

    def carregar_tabela(self, caminho):
        return self.data_importer.carregar_tabela(caminho)

    def obter_colunas(self, tabela):
        return self.data_importer.obter_colunas(tabela)

    def obter_valores_coluna(self, tabela, coluna):
        return self.data_importer.obter_valores_coluna(tabela, coluna)

    @staticmethod
    def sanitizar_nome_arquivo(nome: str, fallback: str) -> str:
        nome_limpo = "".join("_" if c in '\\/:*?"<>|' else c for c in str(nome))
        nome_limpo = nome_limpo.strip().strip(".")
        return nome_limpo or fallback

    @staticmethod
    def normalizar_dado(valor: str, cfg: GeracaoConfig) -> str:
        valor = str(valor)
        if cfg.modo == "numerico":
            return f"{cfg.prefixo}{valor}{cfg.sufixo}"
        return valor

    @staticmethod
    def obter_modelos_barcode():
        return [
            (chave, rotulo)
            for chave, (rotulo, _nome_reportlab) in CodigoService.BARCODE_MODELOS_SUPORTADOS.items()
        ]

    @staticmethod
    def validar_parametros_geracao(codigos, cfg: GeracaoConfig):
        if not isinstance(codigos, list) or not codigos:
            raise ValueError("Nenhum código válido foi encontrado para geração.")
        if len(codigos) > cfg.max_codigos_por_lote:
            raise ValueError(f"Limite excedido: máximo de {cfg.max_codigos_por_lote} códigos por geração.")

        if cfg.qr_width_cm <= 0 or cfg.qr_height_cm <= 0:
            raise ValueError("Tamanho de QR inválido. Informe largura/altura em cm maiores que zero.")
        if cfg.barcode_width_cm <= 0 or cfg.barcode_height_cm <= 0:
            raise ValueError(
                "Tamanho de código de barras inválido. Informe largura/altura em cm maiores que zero."
            )

        if cfg.qr_width_cm > 30 or cfg.qr_height_cm > 30:
            raise ValueError("Tamanho de QR inválido. Use valores até 30 cm.")
        if cfg.barcode_width_cm > 40 or cfg.barcode_height_cm > 20:
            raise ValueError("Tamanho de código de barras inválido. Use até 40x20 cm.")

        validos = []
        invalidos = 0
        for bruto in codigos:
            dado = CodigoService.normalizar_dado(bruto, cfg)
            if not dado or not dado.strip() or len(dado) > cfg.max_tamanho_dado:
                invalidos += 1
                continue
            if cfg.tipo_codigo == "barcode" and any(ord(ch) < 32 for ch in dado):
                invalidos += 1
                continue
            if cfg.tipo_codigo == "barcode":
                try:
                    BarcodeRenderer.validar_modelo(dado.strip(), cfg.barcode_model)
                except ValueError:
                    invalidos += 1
                    continue
            validos.append(str(bruto))

        if not validos:
            raise ValueError("Todos os dados foram rejeitados pela validação de entrada.")
        return validos, invalidos

    def gerar_imagem_obj(self, dado: str, cfg: GeracaoConfig):
        if cfg.tipo_codigo == "barcode":
            return self.barcode_renderer.render(dado, cfg)
        return self.qr_renderer.render(dado, cfg)
