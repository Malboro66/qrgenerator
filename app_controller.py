from dataclasses import dataclass

from app_dependencies import AppDependencies, build_default_dependencies
from models.geracao_config import GeracaoConfig


@dataclass
class AppController:
    deps: AppDependencies

    @classmethod
    def build_default(cls) -> "AppController":
        return cls(deps=build_default_dependencies())

    @property
    def logger(self):
        return self.deps.logger

    @property
    def job_store(self):
        return self.deps.job_store

    @property
    def metrics_store(self):
        return self.deps.metrics_store

    def t(self, key: str, default: str = "", **kwargs) -> str:
        return self.deps.i18n.t(key, default, **kwargs)

    def formatar_excecao(self, exc: Exception, contexto: str) -> str:
        return self.deps.service.formatar_excecao(exc, contexto)

    def carregar_tabela(self, caminho):
        return self.deps.carregar_arquivo_uc.execute(caminho)

    def obter_colunas(self, tabela):
        return self.deps.service.obter_colunas(tabela)

    def obter_valores_coluna(self, tabela, coluna):
        return self.deps.service.obter_valores_coluna(tabela, coluna)

    def validar_parametros_geracao(self, codigos, cfg: GeracaoConfig):
        return self.deps.service.validar_parametros_geracao(codigos, cfg)

    def sanitizar_nome_arquivo(self, nome: str, fallback: str) -> str:
        return self.deps.service.sanitizar_nome_arquivo(nome, fallback)

    def normalizar_dado(self, valor: str, cfg: GeracaoConfig) -> str:
        return self.deps.service.normalizar_dado(valor, cfg)

    def gerar_imagem_obj(self, dado: str, cfg: GeracaoConfig):
        return self.deps.service.gerar_imagem_obj(dado, cfg)

    def extrair_codigos_preview(self, tabela, coluna: str, cfg: GeracaoConfig, max_itens: int):
        return self.deps.atualizar_preview_uc.extrair_codigos_preview(tabela, coluna, cfg, max_itens)

    def gerar_amostra_preview(self, cfg: GeracaoConfig) -> str:
        return self.deps.atualizar_preview_uc.gerar_amostra(cfg)

    def preparar_codigos(self, tabela, coluna: str, cfg: GeracaoConfig):
        return self.deps.gerar_codigos_uc.preparar_codigos(tabela, coluna, cfg)

    def obter_modelos_barcode(self):
        return self.deps.service.obter_modelos_barcode()
