import csv
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from services.data_importer import DataImporter


@pytest.fixture
def imp():
    return DataImporter()


def _csv(dados, encoding="utf-8"):
    fd, caminho = tempfile.mkstemp(suffix=".csv")
    Path(caminho).unlink(missing_ok=True)
    cols = list(dados.keys())
    rows = zip(*[dados[c] for c in cols])
    with open(caminho, "w", newline="", encoding=encoding) as f:
        w = csv.writer(f)
        w.writerow(cols)
        for row in rows:
            w.writerow(row)
    return caminho


def _xlsx(dados):
    pytest.importorskip("openpyxl")
    pd = pytest.importorskip("pandas")
    fd, caminho = tempfile.mkstemp(suffix=".xlsx")
    Path(caminho).unlink(missing_ok=True)
    pd.DataFrame(dados).to_excel(caminho, index=False)
    return caminho


class TestCarregarCSV:
    def test_csv_simples(self, imp):
        caminho = _csv({"id": [1, 2], "nome": ["a", "b"]})
        try:
            tb = imp.carregar_tabela(caminho)
            cols = imp.obter_colunas(tb)
            assert "id" in cols
            assert "nome" in cols
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_csv_utf8_bom(self, imp):
        caminho = _csv({"código": ["A", "B"]}, encoding="utf-8-sig")
        try:
            tb = imp.carregar_tabela(caminho)
            assert "código" in imp.obter_colunas(tb)
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_csv_inexistente_lanca_runtime(self, imp):
        with pytest.raises(RuntimeError, match="Falha ao carregar CSV"):
            imp.carregar_tabela("/tmp/nao_existe_abc123.csv")

    def test_csv_vazio_retorna_lista_vazia_ou_df(self, imp):
        caminho = _csv({"col1": []})
        try:
            tb = imp.carregar_tabela(caminho)
            vals = imp.obter_valores_coluna(tb, "col1")
            assert vals == []
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_csv_fallback_sem_pandas(self, imp):
        caminho = _csv({"col": ["x", "y"]})
        try:
            with patch.dict("sys.modules", {"pandas": None}):
                tb = imp.carregar_tabela(caminho)
            vals = imp.obter_valores_coluna(tb, "col")
            assert vals == ["x", "y"]
        finally:
            Path(caminho).unlink(missing_ok=True)


class TestCarregarXLSX:
    def test_xlsx_simples(self, imp):
        caminho = _xlsx({"sku": ["A1"], "desc": ["item"]})
        try:
            tb = imp.carregar_tabela(caminho)
            assert "sku" in imp.obter_colunas(tb)
        finally:
            Path(caminho).unlink(missing_ok=True)

    def test_xlsx_inexistente_lanca_runtime(self, imp):
        with pytest.raises(RuntimeError, match="Falha ao carregar Excel"):
            imp.carregar_tabela("/tmp/nao_existe_abc123.xlsx")

    def test_xlsx_sem_pandas_lanca_runtime_com_mensagem_clara(self, imp):
        caminho = _xlsx({"a": [1]})
        try:
            with patch.dict("sys.modules", {"pandas": None}):
                with pytest.raises(RuntimeError, match="openpyxl"):
                    imp.carregar_tabela(caminho)
        finally:
            Path(caminho).unlink(missing_ok=True)


class TestObterColunas:
    def test_tabela_none(self, imp):
        assert imp.obter_colunas(None) == []

    def test_dataframe(self, imp):
        pd = pytest.importorskip("pandas")
        tb = pd.DataFrame({"a": [1], "b": [2]})
        assert imp.obter_colunas(tb) == ["a", "b"]

    def test_lista_de_dicts(self, imp):
        assert imp.obter_colunas([{"x": 1, "y": 2}]) == ["x", "y"]

    def test_lista_vazia(self, imp):
        assert imp.obter_colunas([]) == []


class TestObterValoresColuna:
    def test_coluna_com_nulos_ignorados(self, imp):
        pd = pytest.importorskip("pandas")
        tb = pd.DataFrame({"c": ["a", None, "c", float("nan")]})
        assert imp.obter_valores_coluna(tb, "c") == ["a", "c"]

    def test_lista_de_dicts(self, imp):
        tb = [{"k": 1}, {"k": 2}, {"k": None}]
        assert imp.obter_valores_coluna(tb, "k") == ["1", "2"]

    def test_coluna_inexistente_lista(self, imp):
        tb = [{"a": 1}]
        assert imp.obter_valores_coluna(tb, "b") == []


class TestFormatarExcecao:
    def test_mensagem_formatada(self, imp):
        out = imp.formatar_excecao(ValueError("detalhe"), "contexto")
        assert "contexto" in out and "detalhe" in out
