import sys
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "code"))

from pseudocolor_proposed import (  # noqa: E402
    _build_importance_warp,
    _gray_level_structure_importance,
    de_cividis,
    m10_content_adaptive_cividis,
    m11_structure_mass_cividis,
    m12_clahe_cividis,
    m13_local_structure_cividis,
    m14_hybrid_clahe_cividis,
    m15_retinex_cividis,
    m16_unsharp_cividis,
    m17_hist_equalized_cividis,
    m18_bilateral_detail_cividis,
    m19_guided_detail_cividis,
    m20_multiscale_unsharp_cividis,
)


def test_gray_level_importance_marks_structured_gray_range():
    gray = np.full((64, 64), 0.80, dtype=np.float64)
    gray[16:48, 16:48] = 0.35

    importance = _gray_level_structure_importance(gray, bins=64, win=7)

    structured_bin = int(0.35 * 63)
    background_bin = int(0.80 * 63)

    assert importance[structured_bin] > importance[background_bin] * 1.8


def test_importance_warp_allocates_more_lut_distance_to_important_range():
    importance = np.ones(256, dtype=np.float64)
    importance[80:101] = 8.0

    warp = _build_importance_warp(importance, adapt_strength=0.75, smooth_radius=0)

    important_step = np.diff(warp[80:101]).mean()
    ordinary_step = np.diff(warp[20:41]).mean()

    assert np.all(np.diff(warp) > 0)
    assert important_step > ordinary_step * 3.0


def test_content_adaptive_cividis_returns_rgb_float_image():
    gray = np.tile(np.linspace(0, 1, 32, dtype=np.float64), (24, 1))

    out = m10_content_adaptive_cividis(gray)

    assert out.shape == (24, 32, 3)
    assert out.dtype.kind == "f"
    assert out.min() >= 0
    assert out.max() <= 1


def test_new_candidate_methods_return_rgb_float_images():
    gray = np.tile(np.linspace(0, 1, 32, dtype=np.float64), (24, 1))

    for fn in (
        m11_structure_mass_cividis,
        m12_clahe_cividis,
        m13_local_structure_cividis,
        m14_hybrid_clahe_cividis,
        m15_retinex_cividis,
        m16_unsharp_cividis,
        m17_hist_equalized_cividis,
        m18_bilateral_detail_cividis,
        m19_guided_detail_cividis,
        m20_multiscale_unsharp_cividis,
        de_cividis,
    ):
        out = fn(gray)
        assert out.shape == (24, 32, 3)
        assert out.dtype.kind == "f"
        assert out.min() >= 0
        assert out.max() <= 1
