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
from services.print_service import imprimir_png_windows
from services.xml_importer import importar_itens_nfe


class LoteController:
    def __init__(self):
        self.store = LoteStore()
        self.service = LoteService(self.store)

    def importar_xml(self, caminho_xml: str):
        return importar_itens_nfe(caminho_xml)

    def criar_lote(self, seq_lote: int, mes_rec: int, ano_rec: int, nf_ref: str, itens):
        return self.service.criar_lote(seq_lote, mes_rec, ano_rec, nf_ref, itens)

    def listar_lotes(self):
        return self.store.listar_lotes()

    def atualizar_status(self, lote_id: str, status: str):
        self.store.atualizar_status(lote_id, status)

    def deletar_lote(self, lote_id: str):
        self.store.deletar_lote(lote_id)

    def reimprimir_lote(
        self,
        lote: Lote,
        codigos: list[str] | None = None,
        impressora: str = "",
        largura_cm: float = 4.0,
        altura_cm: float = 4.0,
    ):
        codigos_para_imprimir = codigos if codigos is not None else [i.cod_gerado for i in lote.itens]
        return self._imprimir_codigos(codigos_para_imprimir, impressora=impressora, largura_cm=largura_cm, altura_cm=altura_cm)

    def _imprimir_codigos(self, codigos: list[str], impressora: str = "", largura_cm: float = 4.0, altura_cm: float = 4.0):
        ctrl = AppController.build_default()
        cfg = GeracaoConfig(
            qr_width_cm=largura_cm,
            qr_height_cm=altura_cm,
            barcode_width_cm=largura_cm,
            barcode_height_cm=altura_cm,
            keep_qr_ratio=True,
            keep_barcode_ratio=True,
            foreground="black",
            background="white",
            tipo_codigo="barcode",
            barcode_model="code128",
            modo="texto",
            prefixo="",
            sufixo="",
            etiqueta_width_mm=70.0,
            etiqueta_height_mm=50.0,
        )
        validos, _ = ctrl.validar_parametros_geracao(codigos, cfg)
        pasta = Path(tempfile.mkdtemp(prefix="lote_print_"))
        arquivos = []
        for idx, dado in enumerate(validos, 1):
            img = ctrl.gerar_imagem_obj(dado, cfg)
            out = pasta / f"codigo_{idx:03d}.png"
            img.save(str(out), format="PNG", dpi=(ctrl.service.DPI_PADRAO, ctrl.service.DPI_PADRAO))
            arquivos.append(str(out))
            imprimir_png_windows(
                str(out),
                impressora,
                largura_cm,
                altura_cm,
                dpi=ctrl.service.DPI_PADRAO,
                logger=ctrl.logger,
            )

        enviados = True
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
