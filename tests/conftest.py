import os
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def pytest_configure(config):
    config.addinivalue_line("markers", "slow: testes lentos (>2s)")
    config.addinivalue_line("markers", "ui: testes que requerem display tkinter")
    config.addinivalue_line("markers", "windows_only: exclusivos do Windows")


def pytest_collection_modifyitems(config, items):
    if not sys.platform.startswith("win"):
        skip_win = pytest.mark.skip(reason="Requer Windows")
        for item in items:
            if "windows_only" in item.keywords:
                item.add_marker(skip_win)

    sem_display = (not sys.platform.startswith("win")) and (not os.environ.get("DISPLAY"))
    if sem_display:
        skip_ui = pytest.mark.skip(reason="Requer display tkinter ($DISPLAY)")
        for item in items:
            if "ui" in item.keywords:
                item.add_marker(skip_ui)


def pytest_addoption(parser):
    parser.addoption("--slow", action="store_true", default=False)
