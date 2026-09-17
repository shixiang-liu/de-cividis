import sys
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "code"))

from pseudocolor_classic import m5b_turbo, m6_cividis  # noqa: E402


def test_cividis_returns_rgb_float_image():
    gray = np.tile(np.linspace(0, 1, 16, dtype=np.float64), (8, 1))

    out = m6_cividis(gray)

    assert out.shape == (8, 16, 3)
    assert out.dtype.kind == "f"
    assert out.min() >= 0
    assert out.max() <= 1


def test_turbo_returns_rgb_float_image():
    gray = np.tile(np.linspace(0, 1, 16, dtype=np.float64), (8, 1))

    out = m5b_turbo(gray)

    assert out.shape == (8, 16, 3)
    assert out.dtype.kind == "f"
    assert out.min() >= 0
    assert out.max() <= 1
