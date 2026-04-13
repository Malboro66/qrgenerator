import csv

import qrcode
from PIL import Image

from models.geracao_config import GeracaoConfig


class CodigoService:
    """Camada de negócio para carga, validação e geração de códigos."""

    DPI_PADRAO = 200
    BARCODE_MODELOS_SUPORTADOS = {
        "ean13": ("Código de Barras EAN-13", "EAN13"),
        "dun14": ("Código de Barras DUN-14", "I2of5"),
        "upca": ("UPC ou Código Universal de Produto", "UPCA"),
        "code11": ("Code 11", "Code11"),
        "code39": ("Code 39", "Standard39"),
        "code93": ("Code 93", "Standard93"),
        "ean8": ("EAN-8", "EAN8"),
        "interleaved2of5": ("Intercalado 2 de 5", "I2of5"),
        "code128": ("Código 128", "Code128"),
        "gs1128": ("GS1-128", "Code128"),
        "codabar": ("Codabar", "Codabar"),
        "datamatrix": ("Data Matrix", "ECC200DataMatrix"),
    }

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
    def obter_modelos_barcode():
        return [(chave, rotulo) for chave, (rotulo, _nome_reportlab) in CodigoService.BARCODE_MODELOS_SUPORTADOS.items()]

    @staticmethod
    def _validar_modelo_barcode(dado: str, modelo: str):
        if modelo == "ean13" and (not dado.isdigit() or len(dado) not in (12, 13)):
            raise ValueError("EAN-13 exige apenas dígitos com 12 ou 13 caracteres.")
        if modelo == "ean8" and (not dado.isdigit() or len(dado) not in (7, 8)):
            raise ValueError("EAN-8 exige apenas dígitos com 7 ou 8 caracteres.")
        if modelo == "upca" and (not dado.isdigit() or len(dado) not in (11, 12)):
            raise ValueError("UPC-A exige apenas dígitos com 11 ou 12 caracteres.")
        if modelo == "dun14" and (not dado.isdigit() or len(dado) != 14):
            raise ValueError("DUN-14 exige exatamente 14 dígitos numéricos.")
        if modelo == "interleaved2of5":
            if not dado.isdigit():
                raise ValueError("Intercalado 2 de 5 exige apenas dígitos.")
            if len(dado) % 2 != 0:
                raise ValueError("Intercalado 2 de 5 exige quantidade par de dígitos.")

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
            if cfg.tipo_codigo == "barcode":
                try:
                    CodigoService._validar_modelo_barcode(dado.strip(), cfg.barcode_model)
                except ValueError:
                    invalidos += 1
                    continue
            validos.append(str(bruto))

        if not validos:
            raise ValueError('Todos os dados foram rejeitados pela validação de entrada.')
        return validos, invalidos

    @staticmethod
    def _gerar_barcode_pil(dado: str, cfg: GeracaoConfig) -> Image.Image:
        width_px = CodigoService._cm_para_px(cfg.barcode_width_cm)
        height_px = CodigoService._cm_para_px(cfg.barcode_height_cm)
        dado_limpo = dado.strip()
        modelo = cfg.barcode_model or "code128"
        if modelo not in CodigoService.BARCODE_MODELOS_SUPORTADOS:
            raise RuntimeError(f"Modelo de código de barras não suportado: {modelo}")
        CodigoService._validar_modelo_barcode(dado_limpo, modelo)
        _rotulo, nome_reportlab = CodigoService.BARCODE_MODELOS_SUPORTADOS[modelo]

        try:
            from reportlab.graphics import renderPM
            from reportlab.graphics.barcode import createBarcodeDrawing
            from reportlab.lib.units import mm

            opcoes = {"value": dado_limpo}
            if nome_reportlab != "ECC200DataMatrix":
                opcoes.update({"barHeight": 20 * mm, "barWidth": 0.45, "humanReadable": True})
            desenho = createBarcodeDrawing(nome_reportlab, **opcoes)
            img = renderPM.drawToPIL(desenho, dpi=CodigoService.DPI_PADRAO).convert('RGB')
            return CodigoService._resize_with_ratio(img, width_px, height_px, cfg.keep_barcode_ratio)
        except Exception as exc:
            raise RuntimeError(
                "Geração de código de barras indisponível: habilite o backend renderPM do ReportLab."
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
