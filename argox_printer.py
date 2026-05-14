from __future__ import annotations

import logging
import socket
import time
from typing import Any

import serial

logger = logging.getLogger(__name__)


class ArgoxPrinter:
    """Cliente RAW para impressoras Argox (PPLB).

    Exemplo:
        import logging
        logging.basicConfig(level=logging.DEBUG)

        payload = ArgoxPrinter.build_payload(
            text_items=[{"x": 50, "y": 50, "text": "Cód: ABC-123", "font_size": 2, "font_type": 1}],
            barcode_items=[{"x": 50, "y": 150, "data": "1234567890", "height": 60, "type": 1}],
        )

        with ArgoxPrinter(connection_type="network", host="192.168.1.100", port=9100) as printer:
            printer.calibrate_printer(label_length_mm=50, gap_length_mm=3)
            printer.print_label(payload, copies=1)
    """

    def __init__(self, connection_type: str, host: str | None = None, port: int = 9100, serial_port: str | None = None, baudrate: int = 9600, timeout: float = 5.0, encoding: str = "latin-1", sensor_type: str = "r", dpi: int = 203, inter_command_delay: float = 0.05) -> None:
        self.connection_type = connection_type
        self.host = host
        self.port = port
        self.serial_port = serial_port
        self.baudrate = baudrate
        self.timeout = timeout
        self.encoding = encoding
        self.sensor_type = sensor_type
        self.dpi = dpi
        self.inter_command_delay = inter_command_delay
        self._connection: socket.socket | serial.Serial | None = None

    def __enter__(self) -> "ArgoxPrinter":
        """Abre conexão ao entrar no contexto."""
        self.connect()
        return self

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> bool:
        """Fecha conexão ao sair do contexto sem suprimir exceções."""
        self.disconnect()
        return False

    def connect(self) -> bool:
        """Estabelece conexão de rede ou serial."""
        try:
            if self.connection_type == "network":
                if not self.host:
                    raise ValueError("host é obrigatório para conexão de rede")
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(self.timeout)
                sock.connect((self.host, self.port))
                self._connection = sock
            elif self.connection_type == "serial":
                if not self.serial_port:
                    raise ValueError("serial_port é obrigatório para conexão serial")
                self._connection = serial.Serial(self.serial_port, self.baudrate, timeout=self.timeout)
            else:
                raise ValueError("connection_type deve ser 'network' ou 'serial'")
            logger.info("Conexão estabelecida com sucesso")
            return True
        except (socket.error, serial.SerialException, ValueError) as exc:
            logger.error("Falha ao conectar: %s", exc)
            self._connection = None
            return False

    def disconnect(self) -> None:
        """Fecha conexão atual de forma idempotente."""
        conn = self._connection
        self._connection = None
        if conn is None:
            return
        try:
            conn.close()
        except Exception as exc:  # nosec B110
            logger.debug("Erro ao fechar conexão (ignorado): %s", exc)

    def send_command(self, command: str) -> bool:
        """Envia comando bruto para a impressora."""
        if self._connection is None:
            logger.error("Tentativa de envio sem conexão ativa")
            return False
        try:
            payload = command.encode(self.encoding)
            logger.debug("Enviando comando: %r", command)
            if isinstance(self._connection, socket.socket):
                self._connection.sendall(payload)
            else:
                self._connection.write(payload)
            return True
        except (socket.error, serial.SerialException) as exc:
            logger.error("Falha ao enviar comando: %s", exc)
            self.disconnect()
            return False

    def calibrate_printer(self, label_length_mm: float, gap_length_mm: float, stop_position_dots: int | None = None) -> bool:
        """Calibra sensor e tamanho da etiqueta.

        Args:
            label_length_mm: Comprimento da etiqueta em milímetros (> 0).
            gap_length_mm: Gap entre etiquetas em milímetros (> 0).
            stop_position_dots: Posição de parada em dots.
        """
        if label_length_mm <= 0:
            raise ValueError("label_length_mm deve ser maior que zero")
        if gap_length_mm <= 0:
            raise ValueError("gap_length_mm deve ser maior que zero")

        dots_por_mm = self.dpi / 25.4
        label_dots = round(label_length_mm * dots_por_mm)
        gap_dots = round(gap_length_mm * dots_por_mm)
        total_dots = label_dots + gap_dots
        f_value = stop_position_dots if stop_position_dots is not None else max(220, label_dots)

        commands = [f"\x02{self.sensor_type}\r\n", f"Q{total_dots},{gap_dots}\r\n", f"\x02f{f_value}\r\n"]
        for cmd in commands:
            if not self.send_command(cmd):
                return False
            time.sleep(self.inter_command_delay)
        logger.info("Calibração concluída")
        return True

    def print_label(self, pplb_payload: str, copies: int = 1) -> bool:
        """Imprime etiqueta com payload PPLB completo."""
        if copies < 1:
            raise ValueError("copies deve ser >= 1")
        if "E\r\n" not in pplb_payload:
            logger.warning("Payload não contém terminador E\\r\\n")

        sequence = ["N\r\n", pplb_payload, f"P{copies}\r\n"]
        for cmd in sequence:
            if not self.send_command(cmd):
                return False
        logger.info("Etiqueta enviada para impressão")
        return True

    @staticmethod
    def build_payload(text_items: list[dict[str, Any]], barcode_items: list[dict[str, Any]]) -> str:
        """Constrói payload PPLB com comandos de texto e código de barras."""
        lines = ["L\r\n"]
        for item in text_items:
            lines.append(f"H{item['x']},{item['y']},0,{item['font_size']},{item['font_type']},{item['text']}\r\n")
        for item in barcode_items:
            lines.append(f"B{item['x']},{item['y']},0,{item['type']},{item['height']},1,2,{item['data']}\r\n")
        lines.append("E\r\n")
        return "".join(lines)
