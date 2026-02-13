import csv
import io

import qrcode
from PIL import Image

from models.geracao_config import GeracaoConfig


class CodigoService:
    """Camada de negócio para carga, validação e geração de códigos."""

    DPI_PADRAO = 200

    @staticmethod
    def formatar_excecao(exc: Exception, contexto: str) -> str:
        return f"{contexto}: {exc}"

    @staticmethod
    def carregar_tabela(caminho):
        if caminho.lower().endswith('.csv'):
            try:
                import pandas as pd
                return pd.read_csv(caminho)
            except ImportError:
                with open(caminho, newline='', encoding='utf-8-sig') as f:
                    return list(csv.DictReader(f))
            except (OSError, UnicodeDecodeError, ValueError) as exc:
                raise RuntimeError(CodigoService.formatar_excecao(exc, 'Falha ao carregar CSV')) from exc

        try:
            import pandas as pd
            return pd.read_excel(caminho)
        except ImportError as exc:
            raise RuntimeError(
                "Falha ao carregar Excel. Instale/repare 'pandas', 'numpy' e 'openpyxl' no ambiente."
            ) from exc
        except (OSError, ValueError) as exc:
            raise RuntimeError(CodigoService.formatar_excecao(exc, 'Falha ao carregar Excel')) from exc

    @staticmethod
    def obter_colunas(tabela):
        if tabela is None:
            return []
        if hasattr(tabela, 'columns'):
            return list(tabela.columns)
        if isinstance(tabela, list) and tabela:
            return list(tabela[0].keys())
        return []

    @staticmethod
    def obter_valores_coluna(tabela, coluna):
        if hasattr(tabela, '__getitem__') and hasattr(tabela, 'columns'):
            return [str(v) for v in tabela[coluna].dropna().tolist()]
        if isinstance(tabela, list):
            vals = []
            for linha in tabela:
                valor = linha.get(coluna)
                if valor is not None and str(valor).strip() != '':
                    vals.append(str(valor))
            return vals
        return []

    @staticmethod
    def sanitizar_nome_arquivo(nome: str, fallback: str) -> str:
        nome_limpo = ''.join('_' if c in '\\/:*?"<>|' else c for c in str(nome))
        nome_limpo = nome_limpo.strip().strip('.')
        return nome_limpo or fallback

    @staticmethod
    def normalizar_dado(valor: str, cfg: GeracaoConfig) -> str:
        valor = str(valor)
        if cfg.modo == 'numerico':
            return f"{cfg.prefixo}{valor}{cfg.sufixo}"
        return valor

    @staticmethod
    def _cm_para_px(cm: float, dpi: int = DPI_PADRAO) -> int:
        return max(1, int(round((cm / 2.54) * dpi)))

    @staticmethod
    def _resize_with_ratio(img: Image.Image, width_px: int, height_px: int, keep_ratio: bool) -> Image.Image:
        width_px = max(1, width_px)
        height_px = max(1, height_px)
        if not keep_ratio:
            return img.resize((width_px, height_px), Image.Resampling.LANCZOS)

        base = img.copy()
        base.thumbnail((width_px, height_px), Image.Resampling.LANCZOS)
        canvas = Image.new('RGB', (width_px, height_px), 'white')
        x = (width_px - base.width) // 2
        y = (height_px - base.height) // 2
        canvas.paste(base, (x, y))
        return canvas

    @staticmethod
    def validar_parametros_geracao(codigos, cfg: GeracaoConfig):
        if not isinstance(codigos, list) or not codigos:
            raise ValueError('Nenhum código válido foi encontrado para geração.')
        if len(codigos) > cfg.max_codigos_por_lote:
            raise ValueError(f'Limite excedido: máximo de {cfg.max_codigos_por_lote} códigos por geração.')

        if cfg.qr_width_cm <= 0 or cfg.qr_height_cm <= 0:
            raise ValueError('Tamanho de QR inválido. Informe largura/altura em cm maiores que zero.')
        if cfg.barcode_width_cm <= 0 or cfg.barcode_height_cm <= 0:
            raise ValueError('Tamanho de código de barras inválido. Informe largura/altura em cm maiores que zero.')

        # limites práticos
        if cfg.qr_width_cm > 30 or cfg.qr_height_cm > 30:
            raise ValueError('Tamanho de QR inválido. Use valores até 30 cm.')
        if cfg.barcode_width_cm > 40 or cfg.barcode_height_cm > 20:
            raise ValueError('Tamanho de código de barras inválido. Use até 40x20 cm.')

        validos = []
        invalidos = 0
        for bruto in codigos:
            dado = CodigoService.normalizar_dado(bruto, cfg)
            if not dado or not dado.strip() or len(dado) > cfg.max_tamanho_dado:
                invalidos += 1
                continue
            if cfg.tipo_codigo == 'barcode' and any(ord(ch) < 32 for ch in dado):
                invalidos += 1
                continue
            validos.append(str(bruto))

        if not validos:
            raise ValueError('Todos os dados foram rejeitados pela validação de entrada.')
        return validos, invalidos

    @staticmethod
    def _gerar_barcode_pil(dado: str, cfg: GeracaoConfig) -> Image.Image:
        # Backend principal: python-barcode + Pillow (não depende de renderPM).
        width_px = CodigoService._cm_para_px(cfg.barcode_width_cm)
        height_px = CodigoService._cm_para_px(cfg.barcode_height_cm)
        try:
            from barcode import Code128
            from barcode.writer import ImageWriter

            buffer = io.BytesIO()
            writer = ImageWriter()
            codigo = Code128(dado, writer=writer)
            codigo.write(
                buffer,
                options={
                    'module_width': 0.25,
                    'module_height': 18.0,
                    'font_size': 10,
                    'text_distance': 4,
                    'quiet_zone': 2.0,
                    'dpi': CodigoService.DPI_PADRAO,
                },
            )
            buffer.seek(0)
            img = Image.open(buffer).convert('RGB')
            return CodigoService._resize_with_ratio(img, width_px, height_px, cfg.keep_barcode_ratio)
        except (ModuleNotFoundError, ImportError):
            pass
        except Exception as exc:
            raise RuntimeError('Falha ao gerar barcode com python-barcode. Verifique os dados de entrada.') from exc

        # Fallback opcional: renderPM
        try:
            from reportlab.graphics import renderPM
            from reportlab.graphics.barcode import createBarcodeDrawing
            from reportlab.lib.units import mm

            desenho = createBarcodeDrawing('Code128', value=dado, barHeight=20 * mm, barWidth=0.45, humanReadable=True)
            img = renderPM.drawToPIL(desenho, dpi=CodigoService.DPI_PADRAO).convert('RGB')
            return CodigoService._resize_with_ratio(img, width_px, height_px, cfg.keep_barcode_ratio)
        except (ModuleNotFoundError, ImportError, RuntimeError, OSError) as exc:
            raise RuntimeError(
                "Geração de código de barras indisponível: instale a dependência opcional 'python-barcode' "
                "(recomendado) ou habilite o backend renderPM do ReportLab."
            ) from exc

    @staticmethod
    def gerar_imagem_obj(dado: str, cfg: GeracaoConfig) -> Image.Image:
        if cfg.tipo_codigo == 'barcode':
            return CodigoService._gerar_barcode_pil(dado, cfg)

        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(dado)
        qr.make(fit=True)
        img = qr.make_image(fill_color=cfg.foreground, back_color=cfg.background)
        if hasattr(img, 'get_image'):
            img = img.get_image()
        if not isinstance(img, Image.Image):
            img = Image.open(io.BytesIO(img.tobytes()))

        qr_img = img.convert('RGB')
        return CodigoService._resize_with_ratio(
            qr_img,
            CodigoService._cm_para_px(cfg.qr_width_cm),
            CodigoService._cm_para_px(cfg.qr_height_cm),
            cfg.keep_qr_ratio,
        )
