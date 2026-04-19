import json
import sqlite3
from pathlib import Path
from types import SimpleNamespace

import pytest

from application.use_cases import AtualizarPreviewUseCase, CarregarArquivoUseCase, GerarCodigosUseCase
from models.geracao_config import GeracaoConfig
from services.i18n_service import I18nService
from services.job_run_store import JobRunStore
from services.metrics_store import MetricsStore


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
    )
    data.update(kwargs)
    return GeracaoConfig(**data)


class TestCarregarArquivoUseCase:
    def test_delega_para_service(self):
        calls = []

        class S:
            def carregar_tabela(self, caminho):
                calls.append(caminho)
                return [1]

        uc = CarregarArquivoUseCase(service=S())
        out = uc.execute("/tmp/a.csv")
        assert out == [1]
        assert calls == ["/tmp/a.csv"]

    def test_propaga_excecao_do_service(self):
        class S:
            def carregar_tabela(self, caminho):
                raise RuntimeError("boom")

        uc = CarregarArquivoUseCase(service=S())
        with pytest.raises(RuntimeError, match="boom"):
            uc.execute("x")


class TestGerarCodigosUseCase:
    def test_prepara_codigos_chama_validacao(self):
        class S:
            def obter_valores_coluna(self, tabela, coluna):
                return ["a", "b"]

            def validar_parametros_geracao(self, codigos, cfg):
                assert codigos == ["a", "b"]
                return ["a", "b"], 0

        uc = GerarCodigosUseCase(service=S())
        validos, invalidos = uc.preparar_codigos({}, "col", _cfg())
        assert len(validos) == 2 and invalidos == 0

    def test_propaga_valor_invalidos(self):
        class S:
            def obter_valores_coluna(self, tabela, coluna):
                return ["a", ""]

            def validar_parametros_geracao(self, codigos, cfg):
                return ["a"], 1

        uc = GerarCodigosUseCase(service=S())
        validos, invalidos = uc.preparar_codigos({}, "col", _cfg())
        assert validos == ["a"] and invalidos == 1


class TestAtualizarPreviewUseCase:
    def test_extrair_codigos_preview_respeita_max(self):
        class S:
            def obter_valores_coluna(self, tabela, coluna):
                return ["1", "2", "3", "4", "5"]

            def validar_parametros_geracao(self, codigos, cfg):
                return codigos, 0

        uc = AtualizarPreviewUseCase(service=S())
        out = uc.extrair_codigos_preview({}, "id", _cfg(), 3)
        assert len(out) == 3

    def test_extrair_com_max_zero_retorna_vazio(self):
        class S:
            def obter_valores_coluna(self, tabela, coluna):
                return ["1", "2"]

            def validar_parametros_geracao(self, codigos, cfg):
                return codigos, 0

        uc = AtualizarPreviewUseCase(service=S())
        assert uc.extrair_codigos_preview({}, "id", _cfg(), 0) == []

    @pytest.mark.parametrize(
        "modelo,esperado",
        [
            ("code128", "123456789012"),
            ("ean13", "789123456789"),
            ("ean8", "1234567"),
            ("upca", "12345678901"),
            ("dun14", "12345678901231"),
            ("interleaved2of5", "12345678"),
            ("code39", "123456789012"),
            ("code93", "123456789012"),
            ("codabar", "123456789012"),
        ],
    )
    def test_gerar_amostra_barcode(self, modelo, esperado):
        uc = AtualizarPreviewUseCase(service=SimpleNamespace())
        assert uc.gerar_amostra(_cfg(tipo_codigo="barcode", barcode_model=modelo)) == esperado

    def test_gerar_amostra_qrcode(self):
        uc = AtualizarPreviewUseCase(service=SimpleNamespace())
        assert uc.gerar_amostra(_cfg(tipo_codigo="qrcode")) == "https://example.com"


@pytest.fixture
def tmp_locales(tmp_path):
    p = tmp_path / "pt_BR.json"
    e = tmp_path / "en_US.json"
    p.write_text(json.dumps({"hello": "Olá", "greeting": "Olá, {nome}!"}), encoding="utf-8")
    e.write_text(json.dumps({"hello": "Hello", "greeting": "Hello, {nome}!"}), encoding="utf-8")
    return tmp_path


class TestI18nService:
    def test_traduz_chave_existente(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        s.set_locale("pt_BR")
        assert s.t("hello") == "Olá"

    def test_traduz_com_placeholder(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        s.set_locale("pt_BR")
        assert s.t("greeting", nome="João") == "Olá, João!"

    def test_chave_ausente_retorna_default(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        assert s.t("nao_existe", "fallback") == "fallback"

    def test_chave_ausente_sem_default_retorna_chave(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        assert s.t("nao_existe") == "nao_existe"

    def test_troca_locale(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        s.set_locale("en_US")
        assert s.t("hello") == "Hello"

    def test_locale_inexistente_usa_default(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        s.set_locale("de_DE")
        assert s.t("hello") == "Olá"

    def test_json_malformado_retorna_dict_vazio(self, tmp_path):
        (tmp_path / "pt_BR.json").write_text("{", encoding="utf-8")
        s = I18nService(locale_dir=str(tmp_path), default_locale="pt_BR")
        s.set_locale("pt_BR")
        assert s.t("q") == "q"

    def test_placeholder_faltando_retorna_texto_original(self, tmp_locales):
        s = I18nService(locale_dir=str(tmp_locales), default_locale="pt_BR")
        s.set_locale("pt_BR")
        assert "Olá" in s.t("greeting")


@pytest.fixture
def store(tmp_path):
    return JobRunStore(db_path=str(tmp_path / "jobs.db"))


class TestJobRunStore:
    def test_create_run_retorna_uuid(self, store):
        jid = store.create_run(formato="pdf", tipo_codigo="qrcode", modo="texto", destino="x", total_entradas=1, total_invalidos=0)
        assert len(jid) == 36

    def test_update_progress(self, store):
        jid = store.create_run(formato="pdf", tipo_codigo="qrcode", modo="texto", destino="x", total_entradas=3, total_invalidos=0)
        store.update_progress(jid, 3)
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute("SELECT total_processado,status FROM job_runs WHERE id=?", (jid,)).fetchone()
        assert row == (3, "running")

    def test_finish_run_completed(self, store):
        jid = store.create_run(formato="pdf", tipo_codigo="qrcode", modo="texto", destino="x", total_entradas=2, total_invalidos=0)
        store.finish_run(jid, "completed", processado=2)
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute("SELECT status,finished_at,total_processado FROM job_runs WHERE id=?", (jid,)).fetchone()
        assert row[0] == "completed" and row[1] is not None and row[2] == 2

    def test_finish_run_error_com_mensagem(self, store):
        jid = store.create_run(formato="pdf", tipo_codigo="qrcode", modo="texto", destino="x", total_entradas=1, total_invalidos=0)
        store.finish_run(jid, "error", erro="falhou")
        with sqlite3.connect(store.db_path) as conn:
            row = conn.execute("SELECT status,erro FROM job_runs WHERE id=?", (jid,)).fetchone()
        assert row == ("error", "falhou")

    def test_multiplos_jobs_independentes(self, store):
        ids = {
            store.create_run(formato="pdf", tipo_codigo="qrcode", modo="texto", destino="x", total_entradas=1, total_invalidos=0)
            for _ in range(5)
        }
        assert len(ids) == 5

    def test_tabela_criada_automaticamente(self, tmp_path):
        db = tmp_path / "jobs.db"
        JobRunStore(db_path=str(db))
        assert db.exists()


@pytest.fixture
def mstore(tmp_path):
    return MetricsStore(db_path=str(tmp_path / "metrics.db"))


class TestMetricsStore:
    def test_record_run_insere_linha(self, mstore):
        mstore.record_run(formato="pdf", status="completed", total_entradas=1, total_invalidos=0, total_processado=1, duracao_s=1)
        with sqlite3.connect(mstore.db_path) as conn:
            c = conn.execute("SELECT COUNT(*) FROM run_metrics").fetchone()[0]
        assert c == 1

    def test_throughput_calculado(self, mstore):
        mstore.record_run(formato="pdf", status="completed", total_entradas=100, total_invalidos=0, total_processado=100, duracao_s=10)
        snap = mstore.get_health_snapshot()
        assert snap["avg_throughput_itens_s"] == 10.0

    def test_duracao_zero_nao_divide_por_zero(self, mstore):
        mstore.record_run(formato="pdf", status="completed", total_entradas=1, total_invalidos=0, total_processado=10, duracao_s=0)
        snap = mstore.get_health_snapshot()
        assert snap["avg_throughput_itens_s"] == 0.0

    def test_health_snapshot_estrutura(self, mstore):
        snap = mstore.get_health_snapshot()
        for k in ["total_runs", "avg_duration_s", "avg_throughput_itens_s", "error_rate", "by_formato"]:
            assert k in snap

    def test_error_rate_calculada(self, mstore):
        mstore.record_run(formato="pdf", status="completed", total_entradas=1, total_invalidos=0, total_processado=1, duracao_s=1)
        mstore.record_run(formato="pdf", status="error", total_entradas=1, total_invalidos=0, total_processado=0, duracao_s=1)
        snap = mstore.get_health_snapshot()
        assert snap["error_rate"] == 0.5
