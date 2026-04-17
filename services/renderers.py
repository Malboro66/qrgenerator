import io
import qrcode
from PIL import Image, ImageDraw


class ImageResizer:
    @staticmethod
    def resize_with_ratio(img: Image.Image, width_px: int, height_px: int, keep_ratio: bool) -> Image.Image:
        width_px = max(1, width_px)
        height_px = max(1, height_px)
        if not keep_ratio:
            return img.resize((width_px, height_px), Image.Resampling.LANCZOS)

        base = img.copy()
        base.thumbnail((width_px, height_px), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (width_px, height_px), "white")
        x = (width_px - base.width) // 2
        y = (height_px - base.height) // 2
        canvas.paste(base, (x, y))
        return canvas


class QRCodeRenderer:
    def __init__(self, dpi_padrao: int = 200):
        self.dpi_padrao = dpi_padrao

    def _cm_para_px(self, cm: float) -> int:
        return max(1, int(round((cm / 2.54) * self.dpi_padrao)))

    def render(self, dado: str, cfg) -> Image.Image:
        from qrcode.image.svg import SvgImage

        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(dado)
        qr.make(fit=True)
        img = qr.make_image(fill_color=cfg.foreground, back_color=cfg.background)
        if hasattr(img, "get_image"):
            img = img.get_image()

        qr_img = img.convert("RGB")
        return ImageResizer.resize_with_ratio(
            qr_img,
            self._cm_para_px(cfg.qr_width_cm),
            self._cm_para_px(cfg.qr_height_cm),
            cfg.keep_qr_ratio,
        )


# Mapeamento: chave interna → nome no python-barcode
_PYBARCODE_MAP = {
    "code128": "code128",
    "gs1128": "code128",
    "code39": "code39",
    "code93": "code93",
    "ean13": "ean13",
    "ean8": "ean8",
    "upca": "upca",
    "dun14": "itf",
    "interleaved2of5": "itf",
    "codabar": "codabar",
    # datamatrix não tem suporte direto no python-barcode;
    # fica como fallback para reportlab se disponível.
}


class BarcodeRenderer:
    MODELOS_SUPORTADOS = {
        "ean13": ("Código de Barras EAN-13", "EAN13"),
        "dun14": ("Código de Barras DUN-14 (ITF-14)", None),
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

    def __init__(self, dpi_padrao: int = 200):
        self.dpi_padrao = dpi_padrao

    def _cm_para_px(self, cm: float) -> int:
        return max(1, int(round((cm / 2.54) * self.dpi_padrao)))

    @staticmethod
    def validar_modelo(dado: str, modelo: str):
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

    # ------------------------------------------------------------------ #
    #  Backend 1: python-barcode (não precisa de renderPM)                #
    # ------------------------------------------------------------------ #
    def _render_pybarcode(
        self, dado: str, modelo: str, width_px: int, height_px: int, keep_ratio: bool
    ) -> Image.Image:
        import barcode
        from barcode.writer import ImageWriter

        nome_pb = _PYBARCODE_MAP.get(modelo)
        if nome_pb is None:
            raise ValueError(f"Modelo '{modelo}' não suportado pelo python-barcode.")

        bc_class = barcode.get_barcode_class(nome_pb)
        buf = io.BytesIO()
        # options controlam tamanho mínimo; o resize final ajusta para o tamanho pedido
        bc = bc_class(dado, writer=ImageWriter())
        bc.write(buf, options={"write_text": True, "quiet_zone": 2})
        buf.seek(0)
        img = Image.open(buf).convert("RGB")
        img = ImageResizer.resize_with_ratio(img, width_px, height_px, keep_ratio)
        if modelo == "dun14":
            img = self._aplicar_moldura_itf14(img)
        return img

    @staticmethod
    def _aplicar_moldura_itf14(img: Image.Image) -> Image.Image:
        # ITF-14 costuma utilizar "bearer bars" (moldura) para melhorar leitura industrial.
        img_itf14 = img.copy()
        draw = ImageDraw.Draw(img_itf14)
        espessura = max(2, min(img_itf14.width, img_itf14.height) // 40)
        for i in range(espessura):
            draw.rectangle((i, i, img_itf14.width - 1 - i, img_itf14.height - 1 - i), outline="black")
        return img_itf14

    # ------------------------------------------------------------------ #
    #  Backend 2: ReportLab renderPM (original — usado como fallback)     #
    # ------------------------------------------------------------------ #
    def _render_reportlab(
        self, dado: str, modelo: str, width_px: int, height_px: int, keep_ratio: bool
    ) -> Image.Image:
        from reportlab.graphics import renderPM
        from reportlab.graphics.barcode import createBarcodeDrawing
        from reportlab.lib.units import mm as rl_mm

        _rotulo, nome_reportlab = self.MODELOS_SUPORTADOS[modelo]
        if nome_reportlab is None:
            raise RuntimeError("Modelo sem backend reportlab dedicado; use backend ITF-14 via python-barcode.")
        opcoes = {"value": dado}
        if nome_reportlab != "ECC200DataMatrix":
            opcoes.update({"barHeight": 20 * rl_mm, "barWidth": 0.45, "humanReadable": True})
        desenho = createBarcodeDrawing(nome_reportlab, **opcoes)
        img = renderPM.drawToPIL(desenho, dpi=self.dpi_padrao).convert("RGB")
        return ImageResizer.resize_with_ratio(img, width_px, height_px, keep_ratio)

    # ------------------------------------------------------------------ #
    #  Ponto de entrada público                                           #
    # ------------------------------------------------------------------ #
    def render(self, dado: str, cfg) -> Image.Image:
        width_px = self._cm_para_px(cfg.barcode_width_cm)
        height_px = self._cm_para_px(cfg.barcode_height_cm)
        dado_limpo = dado.strip()
        modelo = cfg.barcode_model or "code128"

        if modelo not in self.MODELOS_SUPORTADOS:
            raise RuntimeError(f"Modelo de código de barras não suportado: {modelo}")
        self.validar_modelo(dado_limpo, modelo)

        # Tenta python-barcode primeiro (não requer compilação nativa)
        if modelo in _PYBARCODE_MAP:
            try:
                return self._render_pybarcode(dado_limpo, modelo, width_px, height_px, cfg.keep_barcode_ratio)
            except Exception:
                pass  # fallback abaixo

        # Fallback: reportlab renderPM
        try:
            return self._render_reportlab(dado_limpo, modelo, width_px, height_px, cfg.keep_barcode_ratio)
        except Exception as exc:
            raise RuntimeError(
                "Geração de código de barras indisponível: instale 'python-barcode' "
                "(pip install python-barcode[images]) ou habilite o backend renderPM do ReportLab."
            ) from exc
