"""
tests/test_use_cases.py
========================
Testes para camadas de use-cases, i18n, JobRunStore e MetricsStore.
"""
import json
import os
import sqlite3
import tempfile
import pytest
from unittest.mock import MagicMock, patch

from application.use_cases import (
    AtualizarPreviewUseCase,
    CarregarArquivoUseCase,
    GerarCodigosUseCase,
)
from models.geracao_config import GeracaoConfig
from services.i18n_service import I18nService
from services.job_run_store import JobRunStore
from services.metrics_store import MetricsStore


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


# ─────────────────────────────────────────────
# CarregarArquivoUseCase
# ─────────────────────────────────────────────
class TestCarregarArquivoUseCase:
    def test_delega_para_service(self):
        svc = MagicMock()
        svc.carregar_tabela.return_value = "tabela"
        uc = CarregarArquivoUseCase(svc)
        result = uc.execute("/algum/arquivo.csv")
        svc.carregar_tabela.assert_called_once_with("/algum/arquivo.csv")
        assert result == "tabela"

    def test_propaga_excecao_do_service(self):
        svc = MagicMock()
        svc.carregar_tabela.side_effect = RuntimeError("falha")
        uc = CarregarArquivoUseCase(svc)
        with pytest.raises(RuntimeError, match="falha"):
            uc.execute("/arquivo.csv")


# ─────────────────────────────────────────────
# GerarCodigosUseCase
# ─────────────────────────────────────────────
class TestGerarCodigosUseCase:
    def _svc(self, valores, validos, invalidos):
        svc = MagicMock()
        svc.obter_valores_coluna.return_value = valores
        svc.validar_parametros_geracao.return_value = (validos, invalidos)
        return svc

    def test_prepara_codigos_chama_validacao(self):
        svc = self._svc(["a", "b"], ["a", "b"], 0)
        uc = GerarCodigosUseCase(svc)
        validos, inv = uc.preparar_codigos("tabela", "col", _cfg())
        assert validos == ["a", "b"]
        assert inv == 0

    def test_propaga_valor_invalidos(self):
        svc = self._svc(["a", ""], ["a"], 1)
        uc = GerarCodigosUseCase(svc)
        validos, inv = uc.preparar_codigos("tabela", "col", _cfg())
        assert inv == 1


# ─────────────────────────────────────────────
# AtualizarPreviewUseCase
# ─────────────────────────────────────────────
class TestAtualizarPreviewUseCase:
    def _svc(self, codigos_retorno):
        svc = MagicMock()
        svc.obter_valores_coluna.return_value = codigos_retorno
        svc.validar_parametros_geracao.return_value = (codigos_retorno, 0)
        return svc

    def test_extrair_codigos_preview_respeita_max(self):
        svc = self._svc(["a", "b", "c", "d", "e"])
        uc = AtualizarPreviewUseCase(svc)
        result = uc.extrair_codigos_preview("tabela", "col", _cfg(), max_itens=3)
        assert result == ["a", "b", "c"]

    def test_extrair_com_max_zero_retorna_vazio(self):
        svc = self._svc(["a", "b"])
        uc = AtualizarPreviewUseCase(svc)
        result = uc.extrair_codigos_preview("tabela", "col", _cfg(), max_itens=0)
        assert result == []

    @pytest.mark.parametrize("modelo,esperado", [
        ("ean8",            "1234567"),
        ("ean13",           "789123456789"),
        ("upca",            "12345678901"),
        ("dun14",           "12345678901231"),
        ("interleaved2of5", "12345678"),
        ("code128",         "123456789012"),
    ])
    def test_gerar_amostra_barcode(self, modelo, esperado):
        svc = MagicMock()
        uc = AtualizarPreviewUseCase(svc)
        cfg = _cfg(tipo_codigo="barcode", barcode_model=modelo)
        assert uc.gerar_amostra(cfg) == esperado

    def test_gerar_amostra_qrcode(self):
        svc = MagicMock()
        uc = AtualizarPreviewUseCase(svc)
        result = uc.gerar_amostra(_cfg(tipo_codigo="qrcode"))
        assert result == "https://example.com"


# ─────────────────────────────────────────────
# I18nService
# ─────────────────────────────────────────────
class TestI18nService:
    @pytest.fixture
    def tmp_locales(self, tmp_path):
        pt = {"hello": "Olá", "greeting": "Olá, {nome}!"}
        en = {"hello": "Hello", "greeting": "Hello, {nome}!"}
        (tmp_path / "pt_BR.json").write_text(json.dumps(pt), encoding="utf-8")
        (tmp_path / "en_US.json").write_text(json.dumps(en), encoding="utf-8")
        return str(tmp_path)

    def test_traduz_chave_existente(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        i18n.set_locale("pt_BR")
        assert i18n.t("hello") == "Olá"

    def test_traduz_com_placeholder(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        i18n.set_locale("pt_BR")
        assert i18n.t("greeting", nome="João") == "Olá, João!"

    def test_chave_ausente_retorna_default(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        assert i18n.t("nao_existe", "fallback") == "fallback"

    def test_chave_ausente_sem_default_retorna_chave(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        assert i18n.t("nao_existe") == "nao_existe"

    def test_troca_locale(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        i18n.set_locale("en_US")
        assert i18n.t("hello") == "Hello"

    def test_locale_inexistente_usa_default(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        i18n.set_locale("de_DE")
        assert i18n.t("hello") == "Olá"

    def test_json_malformado_retorna_dict_vazio(self, tmp_path):
        (tmp_path / "pt_BR.json").write_text("{INVALIDO}", encoding="utf-8")
        i18n = I18nService(locale_dir=str(tmp_path), default_locale="pt_BR")
        assert i18n.t("qualquer") == "qualquer"

    def test_placeholder_faltando_retorna_texto_original(self, tmp_locales):
        i18n = I18nService(locale_dir=tmp_locales, default_locale="pt_BR")
        result = i18n.t("greeting")
        assert "Olá" in result


# ─────────────────────────────────────────────
# JobRunStore
# ─────────────────────────────────────────────
class TestJobRunStore:
    @pytest.fixture
    def store(self, tmp_path):
        return JobRunStore(db_path=str(tmp_path / "jobs.db"))

    def test_create_run_retorna_uuid(self, store):
        jid = store.create_run(
            formato="pdf", tipo_codigo="qrcode", modo="texto",
            destino="/tmp/out.pdf", total_entradas=10, total_invalidos=0,
        )
        assert isinstance(jid, str)
        assert len(jid) == 36

    def test_update_progress(self, store):
        jid = store.create_run(
            formato="pdf", tipo_codigo="qrcode", modo="texto",
            destino="/tmp", total_entradas=5, total_invalidos=0,
        )
        store.update_progress(jid, 3)
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute(
                "SELECT total_processado, status FROM job_runs WHERE id=?", (jid,)
            ).fetchone()
        assert row[0] == 3
        assert row[1] == "running"

    def test_finish_run_completed(self, store):
        jid = store.create_run(
            formato="png", tipo_codigo="qrcode", modo="texto",
            destino="/tmp", total_entradas=2, total_invalidos=0,
        )
        store.finish_run(jid, "completed", processado=2)
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute(
                "SELECT status, finished_at, total_processado FROM job_runs WHERE id=?", (jid,)
            ).fetchone()
        assert row[0] == "completed"
        assert row[1] is not None
        assert row[2] == 2

    def test_finish_run_error_com_mensagem(self, store):
        jid = store.create_run(
            formato="pdf", tipo_codigo="barcode", modo="texto",
            destino="/tmp", total_entradas=1, total_invalidos=0,
        )
        store.finish_run(jid, "error", erro="algo deu errado")
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute("SELECT status, erro FROM job_runs WHERE id=?", (jid,)).fetchone()
        assert row[0] == "error"
        assert row[1] == "algo deu errado"

    def test_multiplos_jobs_independentes(self, store):
        ids = [
            store.create_run(
                formato="pdf", tipo_codigo="qrcode", modo="texto",
                destino="/tmp", total_entradas=i, total_invalidos=0,
            )
            for i in range(5)
        ]
        assert len(set(ids)) == 5

    def test_tabela_criada_automaticamente(self, tmp_path):
        db = str(tmp_path / "novo.db")
        s = JobRunStore(db_path=db)
        assert os.path.exists(db)


# ─────────────────────────────────────────────
# MetricsStore
# ─────────────────────────────────────────────
class TestMetricsStore:
    @pytest.fixture
    def store(self, tmp_path):
        return MetricsStore(db_path=str(tmp_path / "metrics.db"))

    def test_record_run_insere_linha(self, store):
        store.record_run(
            formato="pdf", status="completed",
            total_entradas=10, total_invalidos=1,
            total_processado=9, duracao_s=2.5,
        )
        with sqlite3.connect(store.db_path) as conn:
            count = conn.execute("SELECT COUNT(*) FROM run_metrics").fetchone()[0]
        assert count == 1

    def test_throughput_calculado(self, store):
        store.record_run(
            formato="pdf", status="completed",
            total_entradas=100, total_invalidos=0,
            total_processado=100, duracao_s=10.0,
        )
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute(
                "SELECT throughput_itens_s FROM run_metrics"
            ).fetchone()
        assert row[0] == pytest.approx(10.0)

    def test_duracao_zero_nao_divide_por_zero(self, store):
        store.record_run(
            formato="pdf", status="completed",
            total_entradas=0, total_invalidos=0,
            total_processado=0, duracao_s=0.0,
        )
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute(
                "SELECT throughput_itens_s FROM run_metrics"
            ).fetchone()
        assert row[0] == 0.0

    def test_health_snapshot_estrutura(self, store):
        store.record_run(
            formato="pdf", status="completed",
            total_entradas=5, total_invalidos=0,
            total_processado=5, duracao_s=1.0,
        )
        snap = store.get_health_snapshot()
        assert "total_runs" in snap
        assert "avg_duration_s" in snap
        assert "avg_throughput_itens_s" in snap
        assert "error_rate" in snap
        assert "by_formato" in snap
        assert snap["total_runs"] == 1

    def test_error_rate_calculada(self, store):
        store.record_run(formato="pdf", status="completed",
                         total_entradas=1, total_invalidos=0,
                         total_processado=1, duracao_s=1.0)
        store.record_run(formato="pdf", status="error",
                         total_entradas=1, total_invalidos=0,
                         total_processado=0, duracao_s=0.5, erro="falha")
        snap = store.get_health_snapshot()
        assert snap["error_rate"] == pytest.approx(0.5)
