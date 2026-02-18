from __future__ import annotations

from dataclasses import dataclass

from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService


@dataclass(frozen=True)
class CarregarArquivoUseCase:
    service: CodigoService

    def execute(self, caminho: str):
        return self.service.carregar_tabela(caminho)


@dataclass(frozen=True)
class GerarCodigosUseCase:
    service: CodigoService

    def preparar_codigos(self, tabela, coluna: str, cfg: GeracaoConfig):
        codigos = self.service.obter_valores_coluna(tabela, coluna)
        return self.service.validar_parametros_geracao(codigos, cfg)


@dataclass(frozen=True)
class AtualizarPreviewUseCase:
    service: CodigoService

    def extrair_codigos_preview(self, tabela, coluna: str, cfg: GeracaoConfig, max_itens: int):
        codigos = self.service.obter_valores_coluna(tabela, coluna)
        validos, _invalidos = self.service.validar_parametros_geracao(codigos, cfg)
        return validos[: max(0, int(max_itens))]

    def gerar_amostra(self, cfg: GeracaoConfig) -> str:
        if cfg.tipo_codigo == "barcode":
            return "123456789012"
        return "https://example.com"

