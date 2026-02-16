from dataclasses import dataclass
from typing import Any

from models.geracao_config import GeracaoConfig
from services.codigo_service import CodigoService


@dataclass
class CarregarArquivoUseCase:
    service: CodigoService

    def execute(self, caminho: str):
        return self.service.carregar_tabela(caminho)


@dataclass
class GerarCodigosUseCase:
    service: CodigoService

    def preparar_codigos(self, tabela: Any, coluna: str, cfg: GeracaoConfig):
        codigos = self.service.obter_valores_coluna(tabela, coluna)
        return self.service.validar_parametros_geracao(codigos, cfg)


@dataclass
class AtualizarPreviewUseCase:
    service: CodigoService

    def extrair_codigos_preview(self, tabela: Any, coluna: str, cfg: GeracaoConfig, limite: int):
        codigos = self.service.obter_valores_coluna(tabela, coluna)
        codigos, _ = self.service.validar_parametros_geracao(codigos, cfg)
        return codigos[:limite]

    def gerar_amostra(self, cfg: GeracaoConfig):
        dado = self.service.normalizar_dado("123456789", cfg)
        return self.service.gerar_imagem_obj(dado, cfg)
