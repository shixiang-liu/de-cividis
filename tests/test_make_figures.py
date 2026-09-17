import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "code"))

import make_polished_figures  # noqa: E402


def test_fig5_typical_cases_generates_after_baseline_changes(tmp_path, monkeypatch):
    monkeypatch.chdir(ROOT)
    monkeypatch.setattr(make_polished_figures, "SAVE_DPI", 60)

    make_polished_figures.fig5_typical_cases()

    assert os.path.exists(ROOT / "figs" / "fig5_typical_cases.png")
