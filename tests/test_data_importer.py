"""
tests/test_data_importer.py
============================
Testes unitários para DataImporter — carregamento de CSV e Excel.
"""
import csv
import os
import tempfile
import pytest
from unittest.mock import patch, MagicMock
from services.data_importer import DataImporter


@pytest.fixture
def imp():
    return DataImporter()


def _csv(dados: dict, encoding="utf-8") -> str:
    fd, path = tempfile.mkstemp(suffix=".csv")
    with os.fdopen(fd, "w", newline="", encoding=encoding) as f:
        campos = list(dados.keys())
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        n = max(len(v) for v in dados.values())
        for i in range(n):
            w.writerow({k: dados[k][i] for k in campos})
    return path


def _xlsx(dados: dict) -> str:
    try:
        import openpyxl
    except ImportError:
        pytest.skip("openpyxl não instalado")
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb = openpyxl.Workbook()
    ws = wb.active
    campos = list(dados.keys())
    ws.append(campos)
    n = max(len(v) for v in dados.values())
    for i in range(n):
        ws.append([dados[k][i] for k in campos])
    wb.save(path)
    return path


class TestCarregarCSV:
    def test_csv_simples(self, imp):
        path = _csv({"id": [1, 2], "nome": ["a", "b"]})
        try:
            tb = imp.carregar_tabela(path)
            cols = DataImporter.obter_colunas(tb)
            assert "id" in cols
            assert "nome" in cols
        finally:
            os.unlink(path)

    def test_csv_utf8_bom(self, imp):
        path = _csv({"código": ["X1", "X2"]}, encoding="utf-8-sig")
        try:
            tb = imp.carregar_tabela(path)
            cols = DataImporter.obter_colunas(tb)
            assert "código" in cols
        finally:
            os.unlink(path)

    def test_csv_inexistente_lanca_runtime(self, imp):
        with pytest.raises(RuntimeError, match="Falha ao carregar CSV"):
            imp.carregar_tabela("/caminho/que/nao/existe.csv")

    def test_csv_vazio_retorna_lista_vazia_ou_df(self, imp):
        fd, path = tempfile.mkstemp(suffix=".csv")
        try:
            with os.fdopen(fd, "w") as f:
                f.write("col1,col2\n")
            tb = imp.carregar_tabela(path)
            vals = DataImporter.obter_valores_coluna(tb, "col1")
            assert vals == []
        finally:
            os.unlink(path)

    def test_csv_fallback_sem_pandas(self, imp):
        path = _csv({"chave": ["A", "B", "C"]})
        try:
            with patch.dict("sys.modules", {"pandas": None}):
                tb = imp.carregar_tabela(path)
                vals = DataImporter.obter_valores_coluna(tb, "chave")
                assert vals == ["A", "B", "C"]
        finally:
            os.unlink(path)


class TestCarregarXLSX:
    def test_xlsx_simples(self, imp):
        path = _xlsx({"sku": ["001", "002"], "desc": ["Item A", "Item B"]})
        try:
            tb = imp.carregar_tabela(path)
            cols = DataImporter.obter_colunas(tb)
            assert "sku" in cols
        finally:
            os.unlink(path)

    def test_xlsx_inexistente_lanca_runtime(self, imp):
        with pytest.raises(RuntimeError, match="Falha ao carregar Excel"):
            imp.carregar_tabela("/nao/existe.xlsx")

    def test_xlsx_sem_pandas_lanca_runtime_com_mensagem_clara(self, imp):
        with patch.dict("sys.modules", {"pandas": None}):
            with pytest.raises(RuntimeError, match="openpyxl"):
                imp.carregar_tabela("/qualquer.xlsx")


class TestObterColunas:
    def test_tabela_none(self):
        assert DataImporter.obter_colunas(None) == []

    def test_dataframe(self):
        pd = pytest.importorskip("pandas")
        df = pd.DataFrame({"a": [1], "b": [2]})
        assert DataImporter.obter_colunas(df) == ["a", "b"]

    def test_lista_de_dicts(self):
        dados = [{"x": 1, "y": 2}]
        assert DataImporter.obter_colunas(dados) == ["x", "y"]

    def test_lista_vazia(self):
        assert DataImporter.obter_colunas([]) == []


class TestObterValoresColuna:
    def test_coluna_com_nulos_ignorados(self):
        pd = pytest.importorskip("pandas")
        import numpy as np
        df = pd.DataFrame({"v": ["a", None, "c", np.nan]})
        vals = DataImporter.obter_valores_coluna(df, "v")
        assert vals == ["a", "c"]

    def test_lista_de_dicts(self):
        dados = [{"k": "1"}, {"k": "2"}, {"k": None}]
        vals = DataImporter.obter_valores_coluna(dados, "k")
        assert vals == ["1", "2"]

    def test_coluna_inexistente_lista(self):
        dados = [{"a": "x"}]
        vals = DataImporter.obter_valores_coluna(dados, "b")
        assert vals == []


class TestFormatarExcecao:
    def test_mensagem_formatada(self):
        exc = ValueError("detalhe do erro")
        msg = DataImporter.formatar_excecao(exc, "contexto")
        assert "contexto" in msg
        assert "detalhe do erro" in msg
