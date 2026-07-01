import socket
from unittest.mock import MagicMock, patch

import pytest
import serial

from argox_printer import ArgoxPrinter, BarcodeItem, TextItem


class FailingThenSuccessfulSocket(socket.socket):
    def __init__(self):
        super().__init__(socket.AF_INET, socket.SOCK_STREAM)
        self.calls = 0

    def sendall(self, data):
        self.calls += 1
        if self.calls == 1:
            raise socket.error("temporary")


class AlwaysFailingSocket(socket.socket):
    def sendall(self, data):
        raise socket.error("temporary")


def test_connect_network_success():
    with patch("socket.socket") as sock_cls:
        sock = MagicMock()
        sock_cls.return_value = sock
        p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
        assert p.connect() is True


def test_connect_network_failure():
    with patch("socket.socket") as sock_cls:
        sock = MagicMock()
        sock.connect.side_effect = socket.error("x")
        sock_cls.return_value = sock
        p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
        assert p.connect() is False
        assert p._connection is None


def test_connect_invalid_type():
    p = ArgoxPrinter(connection_type="usb")
    assert p.connect() is False


def test_connect_serial_success():
    with patch("serial.Serial"):
        p = ArgoxPrinter(connection_type="serial", serial_port="COM3")
        assert p.connect() is True


def test_disconnect_idempotent():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    p.disconnect()
    p.disconnect()


def test_send_command_encoding_latin1():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    conn = MagicMock()
    p._connection = conn
    assert p.send_command("áéí") is True


def test_send_command_retries_before_disconnect():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1", send_retries=1)
    conn = FailingThenSuccessfulSocket()
    p._connection = conn
    try:
        assert p.send_command("ABC") is True
        assert conn.calls == 2
        assert p._connection is conn
    finally:
        conn.close()


def test_send_command_can_keep_connection_after_error():
    p = ArgoxPrinter(
        connection_type="network",
        host="127.0.0.1",
        disconnect_on_send_error=False,
    )
    conn = AlwaysFailingSocket(socket.AF_INET, socket.SOCK_STREAM)
    p._connection = conn
    try:
        assert p.send_command("ABC") is False
        assert p._connection is conn
    finally:
        conn.close()


def test_send_command_not_connected():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    assert p.send_command("ABC") is False


def test_query_status_network_reads_response():
    class StatusSocket(socket.socket):
        def __init__(self):
            super().__init__(socket.AF_INET, socket.SOCK_STREAM)
            self.sent = b""

        def sendall(self, data):
            self.sent += data

        def recv(self, size):
            assert size == 8
            return b"OK"

    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    conn = StatusSocket()
    p._connection = conn
    try:
        assert p.query_status(read_size=8) == b"OK"
        assert conn.sent == b"\x1b~s"
    finally:
        conn.close()


def test_query_status_invalid_read_size():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    with pytest.raises(ValueError):
        p.query_status(read_size=0)


def test_calibrate_invalid_label_length():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    with pytest.raises(ValueError):
        p.calibrate_printer(50, 0, 1)


def test_calibrate_invalid_label_width():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    with pytest.raises(ValueError):
        p.calibrate_printer(0, 70, 3)


def test_calibrate_sends_correct_sequence():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1", sensor_type="r")
    p.send_command = MagicMock(return_value=True)
    assert p.calibrate_printer(50, 70, 3, stop_position_dots=300) is True
    calls = [c.args[0] for c in p.send_command.call_args_list]
    assert calls[0] == "\x02r\r\n"
    assert calls[1].startswith("q")
    assert calls[2].startswith("Q")
    assert calls[3] == "\x02f300\r\n"
    assert calls[1] == "q400\r\n"
    assert calls[2] == "Q559,24\r\n"


def test_print_label_no_duplicate_L():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    p.send_command = MagicMock(return_value=True)
    payload = "L\r\nH10,10,0,1,1,T\r\nE\r\n"
    p.print_label(payload)
    stream = "".join(c.args[0] for c in p.send_command.call_args_list)
    assert stream.count("L\r\n") == 1


def test_print_label_warns_missing_E(caplog):
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    p.send_command = MagicMock(return_value=True)
    p.print_label("L\r\nH10,10,0,1,1,T\r\n")
    assert "não contém terminador" in caplog.text


def test_print_label_invalid_copies():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    with pytest.raises(ValueError):
        p.print_label("L\r\nE\r\n", copies=0)


def test_context_manager_raises_when_connect_fails():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    p.connect = MagicMock(return_value=False)
    p.disconnect = MagicMock()
    with pytest.raises(ConnectionError, match="127.0.0.1"):
        with p:
            pass
    p.disconnect.assert_not_called()

def test_context_manager_disconnects_on_exit():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    p.connect = MagicMock(return_value=True)
    p.disconnect = MagicMock()
    with pytest.raises(RuntimeError):
        with p:
            raise RuntimeError("boom")
    p.disconnect.assert_called_once()


def test_build_payload_structure():
    payload = ArgoxPrinter.build_payload(
        text_items=[{"x": 50, "y": 50, "text": "Cód", "font_size": 1, "font_type": 1}],
        barcode_items=[{"x": 50, "y": 150, "data": "123", "height": 50, "type": 1}],
    )
    assert payload.startswith("L\r\n")
    assert payload.endswith("E\r\n")


def test_build_payload_with_center_text_and_barcode():
    payload = ArgoxPrinter.build_payload(
        text_items=[{"y": 50, "text": "ABC123", "font_size": 2, "font_type": 1, "center": True, "label_width_dots": 400}],
        barcode_items=[{"y": 150, "data": "123456", "height": 50, "type": 1, "center": True, "label_width_dots": 400}],
    )
    assert "H" in payload
    assert "B" in payload
    assert "H" + str(ArgoxPrinter._calc_center_x(400, ArgoxPrinter._estimate_text_width_dots("ABC123", 2))) in payload


def test_build_payload_uses_manual_width_for_centering():
    payload = ArgoxPrinter.build_payload(
        text_items=[TextItem(text="ABC123", center=True, label_width_dots=400, width_dots=100)],
        barcode_items=[BarcodeItem(data="123456", center=True, label_width_dots=400, width_dots=180)],
    )
    assert "H150,0,0,1,1,ABC123\r\n" in payload
    assert "B110,0,0,1,50,1,2,4,123456\r\n" in payload


def test_build_payload_with_dataclass_defaults():
    payload = ArgoxPrinter.build_payload(text_items=[TextItem(text="ABC")], barcode_items=[BarcodeItem(data="123")])
    assert "H0,0,0,1,1,ABC\r\n" in payload
    assert "B0,0,0,1,50,1,2,4,123\r\n" in payload


def test_build_payload_with_configurable_barcode_options():
    payload = ArgoxPrinter.build_payload(
        text_items=[],
        barcode_items=[BarcodeItem(data="123", human_readable=0, narrow=3, wide=6)],
    )
    assert "B0,0,0,1,50,0,3,6,123\r\n" in payload


def test_build_payload_with_missing_dict_keys_uses_defaults():
    payload = ArgoxPrinter.build_payload(text_items=[{"text": "ABC"}], barcode_items=[{"data": "123"}])
    assert "H0,0,0,1,1,ABC\r\n" in payload
    assert "B0,0,0,1,50,1,2,4,123\r\n" in payload


def test_calc_center_x_invalid():
    with pytest.raises(ValueError):
        ArgoxPrinter._calc_center_x(0, 10)
