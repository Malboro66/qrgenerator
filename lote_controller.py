from __future__ import annotations

import tempfile
from pathlib import Path

from app_controller import AppController
from models.geracao_config import GeracaoConfig
from models.lote import Lote
from services.lote_service import LoteService
from services.lote_store import LoteStore
from services.xml_importer import importar_itens_nfe


class LoteController:
    def __init__(self):
        self.store = LoteStore()
        self.service = LoteService(self.store)

    def importar_xml(self, caminho_xml: str):
        return importar_itens_nfe(caminho_xml)

    def criar_lote(self, mes_rec: int, ano_rec: int, nf_ref: str, ano_nf: str, itens):
        return self.service.criar_lote(mes_rec, ano_rec, nf_ref, ano_nf, itens)

    def listar_lotes(self):
        return self.store.listar_lotes()

    def atualizar_status(self, lote_id: str, status: str):
        self.store.atualizar_status(lote_id, status)

    def reimprimir_lote(self, lote: Lote):
        return self._imprimir_codigos([i.cod_gerado for i in lote.itens])

    def _imprimir_codigos(self, codigos: list[str]):
        ctrl = AppController.build_default()
        cfg = GeracaoConfig(4.0, 4.0, 8.0, 3.0, True, True, "black", "white", "barcode", "code128", "texto", "", "", 100.0, 60.0)
        validos, _invalidos = ctrl.preparar_codigos([{"cod": c} for c in codigos], "cod", cfg)
        pasta = Path(tempfile.mkdtemp(prefix="lote_print_"))
        arquivos = []
        for idx, dado in enumerate(validos, 1):
            img = ctrl.gerar_imagem_obj(dado, cfg)
            out = pasta / f"codigo_{idx:03d}.png"
            img.save(out)
            arquivos.append(str(out))
        return arquivos
