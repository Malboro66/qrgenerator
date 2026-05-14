from __future__ import annotations

import os
import platform
import subprocess
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

    def reimprimir_lote(self, lote: Lote, codigos: list[str] | None = None):
        codigos_para_imprimir = codigos if codigos is not None else [i.cod_gerado for i in lote.itens]
        return self._imprimir_codigos(codigos_para_imprimir)

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

        enviados = self._enviar_para_impressao(arquivos)
        return {"arquivos": arquivos, "enviados": enviados}

    def _enviar_para_impressao(self, arquivos: list[str]) -> bool:
        sistema = platform.system().lower()
        try:
            if "windows" in sistema:
                for arq in arquivos:
                    os.startfile(arq, "print")
                return True
            if "linux" in sistema:
                for arq in arquivos:
                    subprocess.run(["lp", arq], check=False)
                return True
        except Exception:
            return False
        return False
