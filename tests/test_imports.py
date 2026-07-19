"""Smoke tests for the package modules."""

import importlib

import pytest


MODULES = [
    "r7rfa.constants",
    "r7rfa.language",
    "r7rfa.utils",
    "r7rfa.cli",
    "r7rfa.excel_io",
    "r7rfa.chronology",
    "r7rfa.analysis",
    "r7rfa.plots",
    "rugby7_referee_fatigue_analyzer",
]


@pytest.mark.parametrize("module_name", MODULES)
def test_module_can_be_imported(module_name):
    module = importlib.import_module(module_name)

    assert module is not None
