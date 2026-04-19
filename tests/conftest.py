"""
tests/conftest.py
==================
Configuração global do pytest — fixtures compartilhadas e markers.
"""
import os
import sys
import pytest

# Garante que o diretório raiz do projeto esteja no sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def pytest_configure(config):
    config.addinivalue_line("markers", "slow: testes lentos (>2s)")
    config.addinivalue_line("markers", "ui: testes que requerem display tkinter")
    config.addinivalue_line("markers", "windows_only: testes exclusivos do Windows")


def pytest_collection_modifyitems(config, items):
    if not sys.platform.startswith("win"):
        skip_win = pytest.mark.skip(reason="Requer Windows")
        for item in items:
            if "windows_only" in item.keywords:
                item.add_marker(skip_win)

    if os.environ.get("CI"):
        skip_slow = pytest.mark.skip(reason="Pulado em CI (--slow para incluir)")
        for item in items:
            if "slow" in item.keywords and not config.getoption("--slow", default=False):
                item.add_marker(skip_slow)


def pytest_addoption(parser):
    parser.addoption("--slow", action="store_true", default=False,
                     help="Inclui testes marcados como lentos")
