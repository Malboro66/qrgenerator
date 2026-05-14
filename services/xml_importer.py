from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Iterable

from models.lote import ItemLote

NFE_NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}


def _round_qty(value: str) -> int:
    return int(round(float(value)))


def importar_itens_nfe(caminho_xml: str) -> list[ItemLote]:
    try:
        root = ET.parse(caminho_xml).getroot()
    except Exception as exc:
        raise ValueError(f"XML inválido: {exc}") from exc

    inf_nfe = root.find(".//nfe:infNFe", NFE_NS)
    if inf_nfe is None:
        raise ValueError("XML inválido: estrutura NF-e (<infNFe>) não encontrada.")

    n_nf = inf_nfe.findtext("nfe:ide/nfe:nNF", default="", namespaces=NFE_NS).strip()
    if not n_nf:
        raise ValueError("XML inválido: número da NF (<nNF>) não encontrado.")

    itens: list[ItemLote] = []
    for det in inf_nfe.findall("nfe:det", NFE_NS):
        cprod = (det.findtext("nfe:prod/nfe:cProd", default="", namespaces=NFE_NS) or "").strip()
        xprod = (det.findtext("nfe:prod/nfe:xProd", default="", namespaces=NFE_NS) or "").strip()
        qcom = (det.findtext("nfe:prod/nfe:qCom", default="0", namespaces=NFE_NS) or "0").strip()
        try:
            qty = _round_qty(qcom)
        except Exception as exc:
            raise ValueError(f"XML inválido: quantidade inválida no item '{xprod or cprod}'.") from exc
        itens.append(ItemLote(cod_item=cprod, descr_item=xprod, qty=qty, nf_numero=n_nf, cod_gerado=""))

    if not itens:
        raise ValueError("XML inválido: nenhum item <det> encontrado na NF-e.")

    filtrados = [i for i in itens if any(k in i.descr_item.upper() for k in ("INJET", "BOMB"))]
    return filtrados or itens
