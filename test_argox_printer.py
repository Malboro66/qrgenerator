import socket
from unittest.mock import MagicMock, patch

import pytest
import serial

from argox_printer import ArgoxPrinter


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


def test_send_command_not_connected():
    p = ArgoxPrinter(connection_type="network", host="127.0.0.1")
    assert p.send_command("ABC") is False


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


def test_calc_center_x_invalid():
    with pytest.raises(ValueError):
        ArgoxPrinter._calc_center_x(0, 10)
