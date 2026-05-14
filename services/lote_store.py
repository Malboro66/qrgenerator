from __future__ import annotations

import sqlite3
import uuid
from pathlib import Path

from models.lote import ItemLote, Lote


class LoteStore:
    def __init__(self, db_path: str = "logs/lotes.db"):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        with self._connect() as conn:
            conn.execute("""
            CREATE TABLE IF NOT EXISTS lotes (
                id TEXT PRIMARY KEY,
                ident TEXT UNIQUE,
                seq_lote INTEGER,
                mes_rec INTEGER,
                ano_rec INTEGER,
                nf_referencia TEXT,
                criado_em TEXT,
                status TEXT
            )""")
            conn.execute("""
            CREATE TABLE IF NOT EXISTS itens_lote (
                id TEXT PRIMARY KEY,
                lote_id TEXT,
                cod_item TEXT,
                descr_item TEXT,
                qty INTEGER,
                nf_numero TEXT,
                cod_gerado TEXT,
                FOREIGN KEY(lote_id) REFERENCES lotes(id)
            )""")

    def criar_lote(self, lote: Lote) -> str:
        with self._connect() as conn:
            conn.execute("INSERT INTO lotes VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (lote.id, lote.ident, lote.seq_lote, lote.mes_rec, lote.ano_rec, lote.nf_referencia, lote.criado_em, lote.status))
            for item in lote.itens:
                conn.execute("INSERT INTO itens_lote VALUES (?, ?, ?, ?, ?, ?, ?)", (str(uuid.uuid4()), lote.id, item.cod_item, item.descr_item, item.qty, item.nf_numero, item.cod_gerado))
        return lote.id

    def listar_lotes(self) -> list[Lote]:
        with self._connect() as conn:
            rows = conn.execute("SELECT id FROM lotes ORDER BY criado_em DESC").fetchall()
        return [self.obter_lote(r[0]) for r in rows if self.obter_lote(r[0])]

    def obter_lote(self, id: str) -> Lote | None:
        with self._connect() as conn:
            row = conn.execute("SELECT id, ident, seq_lote, mes_rec, ano_rec, nf_referencia, criado_em, status FROM lotes WHERE id=?", (id,)).fetchone()
            if not row:
                return None
            itens_rows = conn.execute("SELECT cod_item, descr_item, qty, nf_numero, cod_gerado FROM itens_lote WHERE lote_id=?", (id,)).fetchall()
        itens = [ItemLote(*r) for r in itens_rows]
        return Lote(id=row[0], ident=row[1], seq_lote=row[2], mes_rec=row[3], ano_rec=row[4], nf_referencia=row[5], criado_em=row[6], status=row[7], itens=itens)

    def atualizar_status(self, id: str, status: str):
        with self._connect() as conn:
            conn.execute("UPDATE lotes SET status=? WHERE id=?", (status, id))

    def proximo_seq_lote(self, mes_rec: int, ano_rec: int) -> int:
        with self._connect() as conn:
            row = conn.execute("SELECT COALESCE(MAX(seq_lote), 0) FROM lotes WHERE mes_rec=? AND ano_rec=?", (mes_rec, ano_rec)).fetchone()
        return int(row[0]) + 1

    def deletar_lote(self, id: str):
        with self._connect() as conn:
            conn.execute("DELETE FROM itens_lote WHERE lote_id=?", (id,))
            conn.execute("DELETE FROM lotes WHERE id=?", (id,))
