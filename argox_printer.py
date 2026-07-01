from __future__ import annotations

import logging
import socket
import time
from dataclasses import dataclass
from typing import Any, ClassVar

import serial

logger = logging.getLogger(__name__)


@dataclass
class TextItem:
    x: int = 0
    y: int = 0
    text: str = ""
    font_size: int = 1
    font_type: int = 1
    center: bool = False
    label_width_dots: int | None = None
    width_dots: int | None = None


@dataclass
class BarcodeItem:
    x: int = 0
    y: int = 0
    data: str = ""
    height: int = 50
    type: int = 1
    human_readable: int = 1
    narrow: int = 2
    wide: int = 4
    center: bool = False
    label_width_dots: int | None = None
    width_dots: int | None = None


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
            printer.calibrate_printer(label_width_mm=50, label_length_mm=50, gap_length_mm=3)
            printer.print_label(payload, copies=1)
    """

    _mm_per_inch: ClassVar[float] = 25.4

    def __init__(self, connection_type: str, host: str | None = None, port: int = 9100, serial_port: str | None = None, baudrate: int = 9600, timeout: float = 5.0, encoding: str = "latin-1", sensor_type: str = "r", dpi: int = 203, inter_command_delay: float = 0.05, send_retries: int = 0, disconnect_on_send_error: bool = True) -> None:
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
        self.send_retries = max(0, send_retries)
        self.disconnect_on_send_error = disconnect_on_send_error
        self._connection: socket.socket | serial.Serial | None = None

    def __enter__(self) -> "ArgoxPrinter":
        """Abre conexão ao entrar no contexto.

        Raises:
            ConnectionError: Quando a conexão não puder ser estabelecida.
        """
        if not self.connect():
            raise ConnectionError(f"Falha ao conectar em {self.host or self.serial_port}")
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

        payload = command.encode(self.encoding)
        max_attempts = self.send_retries + 1
        for attempt in range(1, max_attempts + 1):
            try:
                logger.debug("Enviando comando: %r", command)
                if isinstance(self._connection, socket.socket):
                    self._connection.sendall(payload)
                else:
                    self._connection.write(payload)
                    self._connection.flush()
                return True
            except (socket.error, serial.SerialException) as exc:
                if attempt < max_attempts:
                    logger.warning(
                        "Falha ao enviar comando (tentativa %s/%s): %s",
                        attempt,
                        max_attempts,
                        exc,
                    )
                    time.sleep(self.inter_command_delay)
                    continue
                logger.error("Falha ao enviar comando: %s", exc)
                if self.disconnect_on_send_error:
                    self.disconnect()
                return False
        return False

    def query_status(self, command: str = "\x1b~s", read_size: int = 1024) -> bytes | None:
        """Consulta status da impressora via comando PPLB/ESC.

        O comando padrão ``ESC ~s`` é útil para detectar condições como falta
        de papel quando o modelo/firmware oferece resposta de status.
        """
        if read_size <= 0:
            raise ValueError("read_size deve ser > 0")
        if not self.send_command(command):
            return None
        if self._connection is None:
            logger.error("Não foi possível ler status sem conexão ativa")
            return None
        try:
            if isinstance(self._connection, socket.socket):
                return self._connection.recv(read_size)
            return self._connection.read(read_size)
        except (socket.error, serial.SerialException) as exc:
            logger.error("Falha ao consultar status: %s", exc)
            if self.disconnect_on_send_error:
                self.disconnect()
            return None

    def calibrate_printer(
        self,
        label_width_mm: float,
        label_length_mm: float,
        gap_length_mm: float,
        stop_position_dots: int | None = None,
    ) -> bool:
        """Calibra sensor e tamanho da etiqueta.

        Args:
            label_width_mm: Largura da etiqueta em milímetros (> 0).
            label_length_mm: Comprimento da etiqueta em milímetros (> 0).
            gap_length_mm: Gap entre etiquetas em milímetros (> 0).
            stop_position_dots: Posição de parada em dots.
        """
        if label_width_mm <= 0:
            raise ValueError("label_width_mm deve ser maior que zero")
        if label_length_mm <= 0:
            raise ValueError("label_length_mm deve ser maior que zero")
        if gap_length_mm <= 0:
            raise ValueError("gap_length_mm deve ser maior que zero")

        dots_por_mm = self.dpi / self._mm_per_inch
        width_dots = round(label_width_mm * dots_por_mm)
        label_dots = round(label_length_mm * dots_por_mm)
        gap_dots = round(gap_length_mm * dots_por_mm)
        if width_dots <= 0 or label_dots <= 0 or gap_dots <= 0:
            raise ValueError("Parâmetros convertidos para dots devem ser > 0")
        f_value = stop_position_dots if stop_position_dots is not None else max(220, label_dots)
        if f_value <= 0:
            raise ValueError("stop_position_dots deve ser > 0 quando informado")

        # PPLB:
        # - q define SOMENTE a largura útil da etiqueta (em dots).
        # - Q define SOMENTE comprimento da etiqueta e gap (em dots), nessa ordem.
        # A correção evita o erro antigo de enviar (comprimento + gap) no primeiro parâmetro de Q.
        commands = [
            f"\x02{self.sensor_type}\r\n",
            f"q{width_dots}\r\n",
            f"Q{label_dots},{gap_dots}\r\n",
            f"\x02f{f_value}\r\n",
        ]
        logger.info(
            "Calibração Argox/PPLB: width=%smm(%sdots), length=%smm(%sdots), gap=%smm(%sdots), dpi=%s",
            label_width_mm,
            width_dots,
            label_length_mm,
            label_dots,
            gap_length_mm,
            gap_dots,
            self.dpi,
        )
        for index, cmd in enumerate(commands):
            if not self.send_command(cmd):
                logger.error("Falha ao enviar comando de calibração: %r", cmd)
                return False
            if index < len(commands) - 1:
                time.sleep(self.inter_command_delay)
        logger.info("Calibração concluída")
        return True

    def print_label(self, pplb_payload: str, copies: int = 1) -> bool:
        """Imprime etiqueta com payload PPLB completo."""
        if copies < 1:
            raise ValueError("copies deve ser >= 1")

        payload = pplb_payload
        if "L\r\n" not in payload:
            logger.warning("Payload PPLB não contém início de formato L")
            payload = "L\r\n" + payload
        if "E\r\n" not in payload:
            logger.warning("Payload PPLB não contém terminador E")
            payload = payload + "E\r\n"

        # "N" limpa o buffer, o payload monta/fecha a arte e "P" imprime.
        sequence = ["N\r\n", payload, f"P{copies}\r\n"]
        for index, cmd in enumerate(sequence):
            if not self.send_command(cmd):
                return False
            if index < len(sequence) - 1:
                time.sleep(self.inter_command_delay)
        logger.info("Etiqueta enviada para impressão")
        return True

    @staticmethod
    def _calc_center_x(label_width_dots: int, element_width_dots: int) -> int:
        """Calcula X centralizado na área útil (q) da etiqueta."""
        if label_width_dots <= 0:
            raise ValueError("label_width_dots deve ser > 0")
        if element_width_dots <= 0:
            raise ValueError("element_width_dots deve ser > 0")
        return max(0, (label_width_dots - element_width_dots) // 2)

    @staticmethod
    def _estimate_text_width_dots(text: str, font_size: int) -> int:
        # Heurística PPLB conservadora para centralização horizontal.
        base_char_width = 8
        escala = max(1, int(font_size))
        return max(1, len(text) * base_char_width * escala)

    @staticmethod
    def _estimate_barcode_width_dots(data: str, narrow: int = 2, wide: int = 4) -> int:
        # Heurística aproximada para Code128 (inclui zona de silêncio).
        modulos_aprox = max(1, len(data) * 11 + 35)
        return max(1, modulos_aprox * max(1, narrow) + (wide * 2))

    @staticmethod
    def build_payload(
        text_items: list[TextItem | dict[str, Any]],
        barcode_items: list[BarcodeItem | dict[str, Any]],
    ) -> str:
        """Constrói payload PPLB com comandos de texto e código de barras."""
        lines = ["L\r\n"]
        for raw_item in text_items:
            item = raw_item if isinstance(raw_item, TextItem) else TextItem(**raw_item)
            x = item.x
            if item.center and item.label_width_dots:
                largura = item.width_dots or ArgoxPrinter._estimate_text_width_dots(item.text, item.font_size)
                x = ArgoxPrinter._calc_center_x(item.label_width_dots, largura)
            lines.append(f"H{x},{item.y},0,{item.font_size},{item.font_type},{item.text}\r\n")
        for raw_item in barcode_items:
            item = raw_item if isinstance(raw_item, BarcodeItem) else BarcodeItem(**raw_item)
            x = item.x
            if item.center and item.label_width_dots:
                largura = item.width_dots or ArgoxPrinter._estimate_barcode_width_dots(
                    item.data,
                    narrow=item.narrow,
                    wide=item.wide,
                )
                x = ArgoxPrinter._calc_center_x(item.label_width_dots, largura)
            lines.append(
                f"B{x},{item.y},0,{item.type},{item.height},"
                f"{item.human_readable},{item.narrow},{item.wide},{item.data}\r\n"
            )
        lines.append("E\r\n")
        return "".join(lines)
