from __future__ import annotations

import importlib
import json
import logging
import os
import subprocess
import sys

from PIL import Image


def listar_impressoras_windows() -> tuple[list[str], str]:
    if not sys.platform.startswith("win"):
        return [], ""
    comando = (
        "Get-CimInstance Win32_Printer | "
        "Select-Object Name,Default | "
        "ConvertTo-Json -Compress"
    )
    saida = subprocess.check_output(
        ["powershell", "-NoProfile", "-Command", comando],
        stderr=subprocess.STDOUT,
        timeout=10,
        text=True,
    ).strip()
    if not saida:
        return [], ""

    dados = json.loads(saida)
    if isinstance(dados, dict):
        dados = [dados]

    impressoras = []
    padrao = ""
    for item in dados:
        nome = str(item.get("Name", "")).strip()
        if not nome:
            continue
        impressoras.append(nome)
        if bool(item.get("Default")):
            padrao = nome

    impressoras = sorted(set(impressoras), key=lambda nome: nome.lower())
    return impressoras, padrao


def _imprimir_png_windows_gdi(
    caminho_imagem: str,
    impressora: str,
    largura_cm: float,
    altura_cm: float,
    dpi: int = 200,
    logger=None,
) -> bool:
    logger = logger or logging.getLogger(__name__)
    try:
        win32con = importlib.import_module("win32con")
        win32print = importlib.import_module("win32print")
        win32ui = importlib.import_module("win32ui")
        from PIL import ImageWin
    except Exception:
        return False

    nome_impressora = impressora.strip() if impressora else win32print.GetDefaultPrinter()
    if not nome_impressora:
        return False

    hdc = None
    try:
        img = Image.open(caminho_imagem).convert("RGB")
        img_dpi = img.info.get("dpi", (dpi, dpi))
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(nome_impressora)
        hdc.StartDoc(os.path.basename(caminho_imagem))

        horzsize_mm = max(1, hdc.GetDeviceCaps(win32con.HORZSIZE))
        vertsize_mm = max(1, hdc.GetDeviceCaps(win32con.VERTSIZE))
        horzres_px = max(1, hdc.GetDeviceCaps(win32con.HORZRES))
        vertres_px = max(1, hdc.GetDeviceCaps(win32con.VERTRES))

        ppmm_x = horzres_px / horzsize_mm
        ppmm_y = vertres_px / vertsize_mm

        alvo_w = max(1, int(round(largura_cm * 10.0 * ppmm_x)))
        alvo_h = max(1, int(round(altura_cm * 10.0 * ppmm_y)))

        dib = ImageWin.Dib(img)
        logger.info(
            "Diagnóstico de impressão GDI (ajustado)",
            extra={
                "event": "print_gdi_debug",
                "impressora": nome_impressora,
                "horzsize_mm": horzsize_mm,
                "vertsize_mm": vertsize_mm,
                "horzres_px": horzres_px,
                "vertres_px": vertres_px,
                "ppmm_x": round(ppmm_x, 4),
                "ppmm_y": round(ppmm_y, 4),
                "largura_solicitada_cm": largura_cm,
                "altura_solicitada_cm": altura_cm,
                "alvo_w_px": alvo_w,
                "alvo_h_px": alvo_h,
                "img_original_w_px": img.width,
                "img_original_h_px": img.height,
                "render_dpi_x": img_dpi[0],
                "render_dpi_y": img_dpi[1],
                "alvo_w_mm_calculado": round(alvo_w / ppmm_x, 2) if ppmm_x else 0,
                "alvo_h_mm_calculado": round(alvo_h / ppmm_y, 2) if ppmm_y else 0,
            },
        )

        hdc.StartPage()
        dib.draw(hdc.GetHandleOutput(), (0, 0, alvo_w, alvo_h))
        hdc.EndPage()
        hdc.EndDoc()
        return True
    except Exception:
        return False
    finally:
        if hdc is not None:
            try:
                hdc.DeleteDC()
            except Exception:
                pass


def imprimir_png_windows(
    caminho_imagem: str,
    impressora: str,
    largura_cm: float,
    altura_cm: float,
    dpi: int = 200,
    logger=None,
) -> bool:
    logger = logger or logging.getLogger(__name__)
    if _imprimir_png_windows_gdi(caminho_imagem, impressora, largura_cm, altura_cm, dpi=dpi, logger=logger):
        return True

    if impressora:
        cmd = ["mspaint.exe", "/pt", caminho_imagem, impressora]
    else:
        cmd = ["mspaint.exe", "/p", caminho_imagem]

    try:
        processo = subprocess.Popen(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("Não foi possível localizar o mspaint.exe para realizar a impressão.") from exc
    except OSError as exc:
        raise RuntimeError(f"Falha ao iniciar impressão via mspaint: {exc}") from exc

    if processo.poll() not in (None, 0):
        raise RuntimeError(f"Falha ao enviar imagem para impressão (código {processo.returncode}).")
    return False
