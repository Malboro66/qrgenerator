from dataclasses import dataclass


@dataclass
class GeracaoConfig:
    qr_width_cm: float
    qr_height_cm: float
    barcode_width_cm: float
    barcode_height_cm: float
    keep_qr_ratio: bool
    keep_barcode_ratio: bool
    foreground: str
    background: str
    tipo_codigo: str
    modo: str
    prefixo: str
    sufixo: str
    max_codigos_por_lote: int = 5000
    max_tamanho_dado: int = 512
