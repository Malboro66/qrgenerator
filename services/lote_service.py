from __future__ import annotations

from datetime import datetime
import uuid

from models.lote import Lote, ItemLote
from services.lote_store import LoteStore


class LoteService:
    def __init__(self, store: LoteStore):
        self.store = store

    @staticmethod
    def gerar_ident(seq_lote: int, mes_rec: int, ano_rec: int, nf_ref: str) -> str:
        return f"L{seq_lote:02d}{mes_rec:02d}{str(ano_rec)[-2:]}{str(nf_ref)[-3:]:>03}"

    def proximo_seq_lote(self, mes_rec: int, ano_rec: int) -> int:
        return self.store.proximo_seq_lote(mes_rec, ano_rec)

    def criar_lote(self, seq_lote: int, mes_rec: int, ano_rec: int, nf_referencia: str, itens: list[ItemLote]) -> Lote:
        ident = self.gerar_ident(seq_lote, mes_rec, ano_rec, nf_referencia)
        lote_id = str(uuid.uuid4())
        for i, item in enumerate(itens, start=1):
            item.cod_gerado = f"{ident}_{i:03d}"
        lote = Lote(id=lote_id, ident=ident, seq_lote=seq_lote, mes_rec=mes_rec, ano_rec=int(str(ano_rec)[-2:]), nf_referencia=nf_referencia, itens=itens, criado_em=datetime.now().isoformat(), status="confirmado")
        self.store.criar_lote(lote)
        return lote
