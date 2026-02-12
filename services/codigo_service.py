import csv
import io

import qrcode
from PIL import Image

from models.geracao_config import GeracaoConfig

try:
    from barcode import Code128 as PyCode128
    from barcode.writer import ImageWriter
except Exception:
    PyCode128 = None
    ImageWriter = None


class CodigoService:
    """Camada de negócio para carga, validação e geração de códigos."""

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
    def validar_parametros_geracao(codigos, cfg: GeracaoConfig):
        if not isinstance(codigos, list) or not codigos:
            raise ValueError('Nenhum código válido foi encontrado para geração.')
        if len(codigos) > cfg.max_codigos_por_lote:
            raise ValueError(f'Limite excedido: máximo de {cfg.max_codigos_por_lote} códigos por geração.')
        if cfg.qr_size < 80 or cfg.qr_size > 1200:
            raise ValueError('Tamanho inválido. Use um valor entre 80 e 1200 pixels.')

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
    def _gerar_barcode_pil(dado: str) -> Image.Image:
        # Backend principal: python-barcode + Pillow (não depende de renderPM).
        if PyCode128 and ImageWriter:
            try:
                buffer = io.BytesIO()
                writer = ImageWriter()
                codigo = PyCode128(dado, writer=writer)
                codigo.write(
                    buffer,
                    options={
                        "module_width": 0.25,
                        "module_height": 18.0,
                        "font_size": 10,
                        "text_distance": 4,
                        "quiet_zone": 2.0,
                        "dpi": 200,
                    },
                )
                buffer.seek(0)
                return Image.open(buffer).convert('RGB')
            except Exception as exc:
                raise RuntimeError(
                    "Falha ao gerar barcode com python-barcode. Verifique os dados de entrada."
                ) from exc

        # Fallback opcional: renderPM (mantido por compatibilidade com ambientes legados).
        try:
            from reportlab.graphics import renderPM
            from reportlab.graphics.barcode import createBarcodeDrawing
            from reportlab.graphics.utils import RenderPMError
            from reportlab.lib.units import mm

            desenho = createBarcodeDrawing('Code128', value=dado, barHeight=20 * mm, barWidth=0.45, humanReadable=True)
            return renderPM.drawToPIL(desenho, dpi=200).convert('RGB')
        except (ModuleNotFoundError, ImportError, RuntimeError, OSError) as exc:
            raise RuntimeError(
                "Geração de código de barras indisponível: instale a dependência opcional 'python-barcode' "
                "(recomendado) ou habilite o backend renderPM do ReportLab."
            ) from exc

    @staticmethod
    def gerar_imagem_obj(dado: str, cfg: GeracaoConfig) -> Image.Image:
        if cfg.tipo_codigo == 'barcode':
            return CodigoService._gerar_barcode_pil(dado)
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(dado)
        qr.make(fit=True)
        img = qr.make_image(fill_color=cfg.foreground, back_color=cfg.background)
        if hasattr(img, 'get_image'):
            img = img.get_image()
        if not isinstance(img, Image.Image):
            img = Image.open(io.BytesIO(img.tobytes()))
        return img.convert('RGB')
