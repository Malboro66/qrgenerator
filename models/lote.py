from dataclasses import dataclass, field


@dataclass
class ItemLote:
    cod_item: str
    descr_item: str
    qty: int
    nf_numero: str
    cod_gerado: str = ""


@dataclass
class Lote:
    id: str
    ident: str
    seq_lote: int
    mes_rec: int
    ano_rec: int
    nf_referencia: str
    itens: list[ItemLote] = field(default_factory=list)
    criado_em: str = ""
    status: str = "rascunho"
