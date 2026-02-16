from dataclasses import dataclass


@dataclass
class GeracaoConfig:
    qr_size: int
    barcode_size: int
    foreground: str
    background: str
    tipo_codigo: str
    modo: str
    prefixo: str
    sufixo: str
    max_codigos_por_lote: int = 5000
    max_tamanho_dado: int = 512
