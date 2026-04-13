from dataclasses import dataclass

from application.use_cases import AtualizarPreviewUseCase, CarregarArquivoUseCase, GerarCodigosUseCase
from logging_utils import setup_logging
from services.codigo_service import CodigoService
from services.i18n_service import I18nService
from services.job_run_store import JobRunStore
from services.metrics_store import MetricsStore


@dataclass(frozen=True)
class AppDependencies:
    logger: object
    service: CodigoService
    carregar_arquivo_uc: CarregarArquivoUseCase
    gerar_codigos_uc: GerarCodigosUseCase
    atualizar_preview_uc: AtualizarPreviewUseCase
    job_store: JobRunStore
    metrics_store: MetricsStore
    i18n: I18nService


def build_default_dependencies() -> AppDependencies:
    service = CodigoService()
    return AppDependencies(
        logger=setup_logging(),
        service=service,
        carregar_arquivo_uc=CarregarArquivoUseCase(service),
        gerar_codigos_uc=GerarCodigosUseCase(service),
        atualizar_preview_uc=AtualizarPreviewUseCase(service),
        job_store=JobRunStore(),
        metrics_store=MetricsStore(),
        i18n=I18nService(),
    )
