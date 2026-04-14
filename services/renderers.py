import io

import qrcode
from PIL import Image


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


class BarcodeRenderer:
    MODELOS_SUPORTADOS = {
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

    def render(self, dado: str, cfg) -> Image.Image:
        width_px = self._cm_para_px(cfg.barcode_width_cm)
        height_px = self._cm_para_px(cfg.barcode_height_cm)
        dado_limpo = dado.strip()
        modelo = cfg.barcode_model or "code128"
        if modelo not in self.MODELOS_SUPORTADOS:
            raise RuntimeError(f"Modelo de código de barras não suportado: {modelo}")
        self.validar_modelo(dado_limpo, modelo)
        _rotulo, nome_reportlab = self.MODELOS_SUPORTADOS[modelo]

        try:
            from reportlab.graphics import renderPM
            from reportlab.graphics.barcode import createBarcodeDrawing
            from reportlab.lib.units import mm

            opcoes = {"value": dado_limpo}
            if nome_reportlab != "ECC200DataMatrix":
                opcoes.update({"barHeight": 20 * mm, "barWidth": 0.45, "humanReadable": True})
            desenho = createBarcodeDrawing(nome_reportlab, **opcoes)
            img = renderPM.drawToPIL(desenho, dpi=self.dpi_padrao).convert("RGB")
            return ImageResizer.resize_with_ratio(img, width_px, height_px, cfg.keep_barcode_ratio)
        except Exception as exc:
            try:
                img = self._render_with_python_barcode(dado_limpo, modelo)
                return ImageResizer.resize_with_ratio(img, width_px, height_px, cfg.keep_barcode_ratio)
            except Exception as fallback_exc:
                raise RuntimeError(
                    "Geração de código de barras indisponível: habilite o backend renderPM do ReportLab "
                    "ou instale suporte python-barcode com Pillow."
                ) from fallback_exc

    def _render_with_python_barcode(self, dado_limpo: str, modelo: str) -> Image.Image:
        from barcode import get
        from barcode.writer import ImageWriter

        modelo_python_barcode = {
            "ean13": "ean13",
            "ean8": "ean8",
            "upca": "upc",
            "code39": "code39",
            "code128": "code128",
            "gs1128": "gs1_128",
            "codabar": "codabar",
            "interleaved2of5": "itf",
        }.get(modelo)
        if not modelo_python_barcode:
            raise RuntimeError(f"Modelo de código de barras não suportado sem renderPM: {modelo}")

        barcode_obj = get(modelo_python_barcode, dado_limpo, writer=ImageWriter())
        buffer = io.BytesIO()
        barcode_obj.write(
            buffer,
            options={
                "module_width": 0.3,
                "module_height": 15.0,
                "quiet_zone": 1.0,
                "font_size": 10,
                "text_distance": 3.0,
                "dpi": self.dpi_padrao,
            },
        )
        buffer.seek(0)
        return Image.open(buffer).convert("RGB")
