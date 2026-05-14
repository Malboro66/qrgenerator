from pathlib import Path

import pytest

from services.xml_importer import importar_itens_nfe


def _write_xml(tmp_path: Path, body: str) -> str:
    p = tmp_path / "nfe.xml"
    p.write_text(body, encoding="utf-8")
    return str(p)


def test_xml_valido_retorna_itens(tmp_path):
    xml = """<nfeProc xmlns=\"http://www.portalfiscal.inf.br/nfe\"><NFe><infNFe><ide><nNF>123</nNF></ide><det><prod><cProd>A1</cProd><xProd>INJETOR X</xProd><qCom>2.4</qCom></prod></det></infNFe></NFe></nfeProc>"""
    itens = importar_itens_nfe(_write_xml(tmp_path, xml))
    assert len(itens) == 1
    assert itens[0].qty == 2


def test_filtro_e_fallback(tmp_path):
    xml = """<nfeProc xmlns=\"http://www.portalfiscal.inf.br/nfe\"><NFe><infNFe><ide><nNF>123</nNF></ide><det><prod><cProd>A1</cProd><xProd>parafuso</xProd><qCom>1</qCom></prod></det><det><prod><cProd>A2</cProd><xProd>BOMBA d\'agua</xProd><qCom>3</qCom></prod></det></infNFe></NFe></nfeProc>"""
    itens = importar_itens_nfe(_write_xml(tmp_path, xml))
    assert len(itens) == 1
    xml2 = """<nfeProc xmlns=\"http://www.portalfiscal.inf.br/nfe\"><NFe><infNFe><ide><nNF>123</nNF></ide><det><prod><cProd>A1</cProd><xProd>parafuso</xProd><qCom>1</qCom></prod></det></infNFe></NFe></nfeProc>"""
    itens2 = importar_itens_nfe(_write_xml(tmp_path, xml2))
    assert len(itens2) == 1


def test_xml_invalido(tmp_path):
    p = tmp_path / "x.xml"
    p.write_text("<x></x>", encoding="utf-8")
    with pytest.raises(ValueError):
        importar_itens_nfe(str(p))
