"""
pseudocolor_proposed.py — 候选方法与最终 DE-Cividis 实现

本文件保留早期候选方法以及最终采用的 DE-Cividis。m9_proposed 是
历史融合基线，在最终实验表中记为 M9_old，不再作为本文主方法。
其早期思路包括：
  1) 色盲混淆轴避让（confusion-axis avoidance）：
     在 LMS 色彩空间显式约束伪彩色映射沿安全轴（yellow-blue），
     避免在红绿混淆轴上消耗信息熵。
  2) 直方图自适应映射强度 β(I)：
     根据输入灰度图的直方图统计量 (μ, σ, 熵) 自动选择
     sin-based 的相位与频率参数。
  3) 显著性区域选择性增强：
     用课件 Ch3.7 Sobel + Ch3.3.4 局部统计量构建显著性图，
     对显著区域强化色彩对比，对平滑区域保持感知一致性。

输出：H×W×3 float [0,1]，对色盲友好的伪彩色图像。
"""
import numpy as np
import cv2
import matplotlib.cm as cm

from cvd_simulate import simulate_cvd, RGB2LMS
from saliency import saliency_map
from pseudocolor_classic import m2_sin_three_phase, m6_viridis, _to_float01


def _histogram_statistics(gray):
    """计算图像直方图的均值、标准差、归一化熵。

    Returns:
        dict with keys 'mean', 'std', 'entropy'
    """
    g = _to_float01(gray)
    mean = float(g.mean())
    std = float(g.std())
    # 计算熵
    hist, _ = np.histogram(g, bins=256, range=(0, 1))
    p = hist.astype(np.float64) / max(hist.sum(), 1)
    p_nz = p[p > 0]
    entropy = float(-(p_nz * np.log2(p_nz)).sum())  # bits, max=8
    return {"mean": mean, "std": std, "entropy": entropy / 8.0}


def _normalize01(arr):
    """Return arr normalized to [0, 1]."""
    x = np.asarray(arr, dtype=np.float64)
    lo = float(np.min(x))
    hi = float(np.max(x))
    if hi <= lo + 1e-12:
        return np.zeros_like(x, dtype=np.float64)
    return (x - lo) / (hi - lo)


def _gray_level_structure_importance(gray, bins=256, win=15, edge_weight=0.6):
    """
    Estimate how much structural information each gray-level interval carries.

    This is the profiling step for the adaptive LUT: pixels with strong Sobel
    edges or high local variance vote for the gray-level interval they belong
    to. The result is a 1-D importance curve over gray values.
    """
    g = _to_float01(gray).astype(np.float32)

    gx = cv2.Sobel(g, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(g, cv2.CV_32F, 0, 1, ksize=3)
    edge = _normalize01(np.sqrt(gx * gx + gy * gy))

    kernel = (win, win)
    mean = cv2.boxFilter(g, cv2.CV_32F, kernel)
    mean_sq = cv2.boxFilter(g * g, cv2.CV_32F, kernel)
    local_std = _normalize01(np.sqrt(np.maximum(mean_sq - mean * mean, 0.0)))

    structure = edge_weight * edge + (1.0 - edge_weight) * local_std
    idx = np.clip((g * (bins - 1)).astype(np.int32), 0, bins - 1)

    sums = np.bincount(idx.ravel(), weights=structure.ravel(), minlength=bins)
    counts = np.bincount(idx.ravel(), minlength=bins).astype(np.float64)
    importance = np.zeros(bins, dtype=np.float64)
    np.divide(sums, counts, out=importance, where=counts > 0)
    return importance


def _structure_response(gray, win=15, edge_weight=0.6):
    """Compute a lightweight structure response from edges and local variance."""
    g = _to_float01(gray).astype(np.float32)
    gx = cv2.Sobel(g, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(g, cv2.CV_32F, 0, 1, ksize=3)
    edge = _normalize01(np.sqrt(gx * gx + gy * gy))

    kernel = (win, win)
    mean = cv2.boxFilter(g, cv2.CV_32F, kernel)
    mean_sq = cv2.boxFilter(g * g, cv2.CV_32F, kernel)
    local_std = _normalize01(np.sqrt(np.maximum(mean_sq - mean * mean, 0.0)))

    return np.clip(edge_weight * edge + (1.0 - edge_weight) * local_std, 0.0, 1.0)


def _gray_level_structure_mass(gray, bins=256, win=15, edge_weight=0.6):
    """
    Structure-weighted gray histogram.

    Unlike _gray_level_structure_importance, this keeps pixel support in the
    vote. A rare noisy gray level cannot dominate merely because its average
    response is high.
    """
    g = _to_float01(gray).astype(np.float32)
    structure = _structure_response(g, win=win, edge_weight=edge_weight)
    idx = np.clip((g * (bins - 1)).astype(np.int32), 0, bins - 1)
    mass = np.bincount(idx.ravel(), weights=structure.ravel(), minlength=bins)
    return mass.astype(np.float64)


def _build_importance_warp(importance, adapt_strength=0.65, smooth_radius=3):
    """
    Build a monotone gray->colormap-position warp from gray-level importance.

    Larger importance means larger increments in the resulting mapping, so the
    corresponding gray interval receives more perceptual color distance.
    """
    imp = np.maximum(np.asarray(importance, dtype=np.float64), 0.0)
    if smooth_radius > 0:
        radius = int(smooth_radius)
        kernel = np.ones(2 * radius + 1, dtype=np.float64)
        kernel /= kernel.sum()
        imp = np.convolve(np.pad(imp, radius, mode="edge"), kernel, mode="valid")

    imp_norm = _normalize01(imp)
    strength = float(np.clip(adapt_strength, 0.0, 0.95))
    density = (1.0 - strength) + strength * imp_norm
    density = np.maximum(density, 1e-6)

    cumulative = np.concatenate([[0.0], np.cumsum(density[:-1])])
    if cumulative[-1] <= 0:
        return np.linspace(0.0, 1.0, len(imp), dtype=np.float64)
    warp = cumulative / cumulative[-1]
    return np.clip(warp, 0.0, 1.0)


def _apply_ycbcr_luminance_constraint(rgb, gray):
    """Replace output luminance with the source gray image."""
    rgb_clamped = np.clip(rgb, 0, 1)
    g = _to_float01(gray)
    return _replace_ycbcr_luminance(rgb_clamped, g)


def _replace_ycbcr_luminance(rgb, luminance, chroma_scale=1.0):
    """Replace Y in a simple YCbCr transform and optionally scale chroma."""
    rgb_clamped = np.clip(rgb, 0, 1)
    y = _to_float01(luminance)
    chroma = float(np.clip(chroma_scale, 0.0, 2.0))
    cb = -0.169 * rgb_clamped[..., 0] - 0.331 * rgb_clamped[..., 1] + 0.500 * rgb_clamped[..., 2] + 0.5
    cr = 0.500 * rgb_clamped[..., 0] - 0.419 * rgb_clamped[..., 1] - 0.081 * rgb_clamped[..., 2] + 0.5
    cb = 0.5 + chroma * (cb - 0.5)
    cr = 0.5 + chroma * (cr - 0.5)
    r = y + 1.402 * (cr - 0.5)
    gg = y - 0.344 * (cb - 0.5) - 0.714 * (cr - 0.5)
    b = y + 1.772 * (cb - 0.5)
    return np.clip(np.stack([r, gg, b], axis=-1), 0.0, 1.0)


def m10_content_adaptive_cividis(
    rgb_or_gray,
    bins=256,
    win=15,
    edge_weight=0.6,
    adapt_strength=0.65,
    smooth_radius=3,
    luminance_constraint=True,
):
    """
    Content-adaptive CVD-friendly pseudocolor mapping.

    The method uses cividis as a CVD-safe base colormap, but re-parameterizes
    it per image. Gray ranges that carry more edge/texture structure receive
    larger colormap distance; flat background ranges are compressed.
    """
    arr = np.asarray(rgb_or_gray)
    if arr.ndim == 3 and arr.shape[2] == 3:
        if arr.dtype == np.uint8:
            gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0
        else:
            arr_u = (np.clip(arr, 0, 1) * 255).astype(np.uint8)
            gray = cv2.cvtColor(arr_u, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0
    else:
        gray = _to_float01(arr)

    importance = _gray_level_structure_importance(
        gray, bins=bins, win=win, edge_weight=edge_weight
    )
    warp = _build_importance_warp(
        importance, adapt_strength=adapt_strength, smooth_radius=smooth_radius
    )

    idx = np.clip((gray * (bins - 1)).astype(np.int32), 0, bins - 1)
    warped_gray = warp[idx]
    rgb = cm.cividis(warped_gray)[..., :3]

    if luminance_constraint:
        rgb = _apply_ycbcr_luminance_constraint(rgb, gray)
    return np.clip(rgb, 0.0, 1.0)


def m11_structure_mass_cividis(
    rgb_or_gray,
    bins=256,
    win=15,
    edge_weight=0.6,
    adapt_strength=0.65,
    smooth_radius=3,
    luminance_constraint=True,
):
    """
    Candidate: structure-mass weighted cividis reparameterization.

    Uses the amount of structural response per gray interval, not just the
    average response. This is less sensitive to sparse noisy gray levels.
    """
    gray = _to_float01(rgb_or_gray)
    mass = _gray_level_structure_mass(gray, bins=bins, win=win, edge_weight=edge_weight)
    warp = _build_importance_warp(
        mass, adapt_strength=adapt_strength, smooth_radius=smooth_radius
    )
    idx = np.clip((gray * (bins - 1)).astype(np.int32), 0, bins - 1)
    rgb = cm.cividis(warp[idx])[..., :3]
    if luminance_constraint:
        rgb = _apply_ycbcr_luminance_constraint(rgb, gray)
    return np.clip(rgb, 0.0, 1.0)


def m12_clahe_cividis(
    rgb_or_gray,
    clip_limit=2.0,
    tile_grid_size=(8, 8),
    luminance_constraint=True,
):
    """
    Candidate: CLAHE-preconditioned cividis.

    Local contrast enhancement is applied to gray values before the CVD-safe
    colormap. This is a strong classical baseline for structure visibility.
    """
    gray = _to_float01(rgb_or_gray)
    u8 = np.clip(gray * 255.0, 0, 255).astype(np.uint8)
    clahe = cv2.createCLAHE(clipLimit=float(clip_limit), tileGridSize=tile_grid_size)
    enhanced = clahe.apply(u8).astype(np.float64) / 255.0
    rgb = cm.cividis(enhanced)[..., :3]
    if luminance_constraint:
        rgb = _apply_ycbcr_luminance_constraint(rgb, gray)
    return np.clip(rgb, 0.0, 1.0)


def m13_local_structure_cividis(
    rgb_or_gray,
    win=15,
    edge_weight=0.6,
    boost=0.35,
    luminance_constraint=True,
):
    """
    Candidate: local structure boosted cividis.

    This is not a pure 1-D LUT: structure response locally increases contrast
    before mapping. It is included as an upper-bound classical baseline.
    """
    gray = _to_float01(rgb_or_gray)
    g = gray.astype(np.float32)
    mean = cv2.boxFilter(g, cv2.CV_32F, (win, win)).astype(np.float64)
    structure = _structure_response(gray, win=win, edge_weight=edge_weight)
    enhanced = np.clip(gray + boost * structure * (gray - mean), 0.0, 1.0)
    rgb = cm.cividis(enhanced)[..., :3]
    if luminance_constraint:
        rgb = _apply_ycbcr_luminance_constraint(rgb, gray)
    return np.clip(rgb, 0.0, 1.0)


def m14_hybrid_clahe_cividis(
    rgb_or_gray,
    clip_limit=4.0,
    tile_grid_size=(8, 8),
    clahe_alpha=1.0,
    gamma=1.0,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: tunable CLAHE-cividis hybrid.

    This generalizes M12. CLAHE can be blended with the original gray values,
    gamma can redistribute the enhanced gray levels, luminance can be mixed
    between original gray and enhanced gray, and chroma can be scaled.
    """
    gray = _to_float01(rgb_or_gray)
    u8 = np.clip(gray * 255.0, 0, 255).astype(np.uint8)
    clahe = cv2.createCLAHE(clipLimit=float(clip_limit), tileGridSize=tile_grid_size)
    enhanced = clahe.apply(u8).astype(np.float64) / 255.0

    alpha = float(np.clip(clahe_alpha, 0.0, 1.0))
    mapped_gray = np.clip((1.0 - alpha) * gray + alpha * enhanced, 0.0, 1.0)

    gam = max(float(gamma), 1e-3)
    mapped_gray = np.clip(mapped_gray, 0.0, 1.0) ** gam
    rgb = cm.cividis(mapped_gray)[..., :3]

    lum_mix = float(np.clip(luminance_mix, 0.0, 1.0))
    target_y = np.clip(lum_mix * gray + (1.0 - lum_mix) * mapped_gray, 0.0, 1.0)
    rgb = _replace_ycbcr_luminance(rgb, target_y, chroma_scale=chroma_scale)
    return np.clip(rgb, 0.0, 1.0)


def _map_cividis_with_luminance(gray, mapped_gray, luminance_mix=1.0, chroma_scale=1.0):
    """Map a gray transform through cividis and restore/mix luminance."""
    rgb = cm.cividis(np.clip(mapped_gray, 0.0, 1.0))[..., :3]
    lum_mix = float(np.clip(luminance_mix, 0.0, 1.0))
    target_y = np.clip(lum_mix * gray + (1.0 - lum_mix) * mapped_gray, 0.0, 1.0)
    return _replace_ycbcr_luminance(rgb, target_y, chroma_scale=chroma_scale)


def m15_retinex_cividis(
    rgb_or_gray,
    sigmas=(15, 80),
    retinex_alpha=0.65,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: multi-scale Retinex-style gray enhancement + cividis.

    Retinex emphasizes reflectance-like local changes by subtracting blurred
    log intensity at several scales. This is a different family from CLAHE.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float64)
    g = np.clip(gray, 1e-4, 1.0)
    accum = np.zeros_like(g, dtype=np.float64)
    for sigma in sigmas:
        blur = cv2.GaussianBlur(g, (0, 0), sigmaX=float(sigma), sigmaY=float(sigma))
        accum += np.log(g) - np.log(np.clip(blur, 1e-4, 1.0))
    ret = _normalize01(accum / max(len(sigmas), 1))
    alpha = float(np.clip(retinex_alpha, 0.0, 1.0))
    mapped = np.clip((1.0 - alpha) * gray + alpha * ret, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


def m16_unsharp_cividis(
    rgb_or_gray,
    sigma=2.0,
    amount=1.2,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: unsharp/detail-boosted cividis.

    This tests whether explicit spatial detail enhancement beats histogram-
    based enhancement while keeping the same CVD-safe base colormap.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float64)
    blur = cv2.GaussianBlur(gray, (0, 0), sigmaX=float(sigma), sigmaY=float(sigma))
    detail = gray - blur
    mapped = np.clip(gray + float(amount) * detail, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


def m17_hist_equalized_cividis(
    rgb_or_gray,
    equalize_alpha=1.0,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: global histogram equalization + cividis.

    This is a global contrast baseline, intentionally separate from CLAHE.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float64)
    u8 = np.clip(gray * 255.0, 0, 255).astype(np.uint8)
    eq = cv2.equalizeHist(u8).astype(np.float64) / 255.0
    alpha = float(np.clip(equalize_alpha, 0.0, 1.0))
    mapped = np.clip((1.0 - alpha) * gray + alpha * eq, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


def m18_bilateral_detail_cividis(
    rgb_or_gray,
    diameter=9,
    sigma_color=0.08,
    sigma_space=5.0,
    amount=2.0,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: bilateral-filter detail boost + cividis.

    Bilateral filtering estimates an edge-preserving base layer. Subtracting it
    gives a detail layer that can be boosted without smearing strong edges as
    much as Gaussian unsharp masking.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float32)
    d = int(max(3, diameter))
    if d % 2 == 0:
        d += 1
    base = cv2.bilateralFilter(
        gray,
        d=d,
        sigmaColor=float(sigma_color),
        sigmaSpace=float(sigma_space),
    ).astype(np.float64)
    detail = gray.astype(np.float64) - base
    mapped = np.clip(gray.astype(np.float64) + float(amount) * detail, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


def _guided_filter_gray(guide, src, radius=8, eps=1e-3):
    """Guided filter for single-channel float images."""
    i = _to_float01(guide).astype(np.float32)
    p = _to_float01(src).astype(np.float32)
    r = int(max(1, radius))
    ksize = (2 * r + 1, 2 * r + 1)

    mean_i = cv2.boxFilter(i, cv2.CV_32F, ksize)
    mean_p = cv2.boxFilter(p, cv2.CV_32F, ksize)
    corr_i = cv2.boxFilter(i * i, cv2.CV_32F, ksize)
    corr_ip = cv2.boxFilter(i * p, cv2.CV_32F, ksize)

    var_i = corr_i - mean_i * mean_i
    cov_ip = corr_ip - mean_i * mean_p

    a = cov_ip / (var_i + float(eps))
    b = mean_p - a * mean_i
    mean_a = cv2.boxFilter(a, cv2.CV_32F, ksize)
    mean_b = cv2.boxFilter(b, cv2.CV_32F, ksize)
    return (mean_a * i + mean_b).astype(np.float64)


def m19_guided_detail_cividis(
    rgb_or_gray,
    radius=8,
    eps=1e-3,
    amount=2.0,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: guided-filter detail boost + cividis.

    Guided filtering creates an edge-aware base layer using the original gray
    image as guidance, then boosts the residual detail layer.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float64)
    base = _guided_filter_gray(gray, gray, radius=radius, eps=eps)
    detail = gray - base
    mapped = np.clip(gray + float(amount) * detail, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


def m20_multiscale_unsharp_cividis(
    rgb_or_gray,
    sigmas=(1.0, 3.0, 8.0),
    weights=(0.5, 0.35, 0.15),
    amount=2.0,
    luminance_mix=1.0,
    chroma_scale=1.0,
):
    """
    Candidate: multi-scale unsharp/detail boost + cividis.

    Instead of using one Gaussian scale, this combines residuals from several
    blur scales so both fine and medium structures can receive color distance.
    """
    gray = _to_float01(rgb_or_gray).astype(np.float64)
    sig = tuple(float(s) for s in sigmas)
    w = np.asarray(weights, dtype=np.float64)
    if len(w) != len(sig):
        w = np.ones(len(sig), dtype=np.float64)
    w = w / max(float(w.sum()), 1e-8)

    detail = np.zeros_like(gray, dtype=np.float64)
    for sigma, weight in zip(sig, w):
        blur = cv2.GaussianBlur(gray, (0, 0), sigmaX=sigma, sigmaY=sigma)
        detail += weight * (gray - blur)
    mapped = np.clip(gray + float(amount) * detail, 0.0, 1.0)
    return _map_cividis_with_luminance(gray, mapped, luminance_mix, chroma_scale)


DE_LUMINANCE_MIX = 0.9


def de_cividis(rgb_or_gray):
    """
    Paper-facing alias for the final method.

    DE-Cividis means Detail-Enhanced Cividis: Gaussian detail extraction,
    unsharp enhancement with sigma=2.0 and amount=2.0, cividis mapping, and
    luminance write-back dominated by the original gray image.
    """
    return m16_unsharp_cividis(
        rgb_or_gray,
        sigma=2.0,
        amount=2.0,
        luminance_mix=DE_LUMINANCE_MIX,
        chroma_scale=1.0,
    )


def de_cividis_no_de(rgb_or_gray):
    """Ablation variant: skip detail enhancement, equivalent to fixed Cividis."""
    gray = _to_float01(rgb_or_gray)
    return _map_cividis_with_luminance(
        gray,
        gray,
        luminance_mix=DE_LUMINANCE_MIX,
        chroma_scale=1.0,
    )


def de_cividis_no_lw(rgb_or_gray):
    """Ablation variant: keep detail enhancement but skip YCbCr luminance write-back."""
    gray = _to_float01(rgb_or_gray)
    blur = cv2.GaussianBlur(gray, (0, 0), sigmaX=2.0, sigmaY=2.0)
    detail = gray - blur
    mapped = np.clip(gray + 2.0 * detail, 0.0, 1.0)
    return np.clip(cm.cividis(mapped)[..., :3], 0.0, 1.0)


def de_jet_no_cividis(rgb_or_gray):
    """Ablation variant: use Jet instead of Cividis after detail enhancement."""
    gray = _to_float01(rgb_or_gray)
    blur = cv2.GaussianBlur(gray, (0, 0), sigmaX=2.0, sigmaY=2.0)
    detail = gray - blur
    mapped = np.clip(gray + 2.0 * detail, 0.0, 1.0)
    rgb = cm.jet(mapped)[..., :3]
    target_y = np.clip(
        DE_LUMINANCE_MIX * gray + (1.0 - DE_LUMINANCE_MIX) * mapped,
        0.0,
        1.0,
    )
    return _replace_ycbcr_luminance(rgb, target_y)


def _adaptive_sin_params(stats):
    """
    创新点 2：根据直方图统计量自适应选择 sin-based 的频率和相位。

    设计逻辑：
    - 高熵图（信息丰富）：用低频映射，让相邻灰度值的颜色变化平滑
    - 低熵图（对比强烈）：用稍高频，强化对比度
    - 暗图：相位偏向蓝紫
    - 亮图：相位偏向黄绿

    Returns:
        (freq, phase_r, phase_g, phase_b)
    """
    e = stats["entropy"]  # [0, 1]
    m = stats["mean"]     # [0, 1]

    # 频率：低熵 → 频率高（增强对比），高熵 → 频率低（平滑过渡）
    freq = 1.0 - 0.4 * e  # 范围约 [0.6, 1.0]

    # 相位偏移：根据均值调整起点
    # 暗图 → 起点蓝紫（相位偏后）
    # 亮图 → 起点蓝紫（相位偏前）
    phase_offset = (m - 0.5) * np.pi * 0.5  # [-π/4, π/4]

    # 创新点 1 关键：避让红绿混淆轴
    # sin-based 的三相位默认是 0, 2π/3, 4π/3，导致 R 和 G 通道在不同 I 时相反
    # 改为：让 R 和 G 通道相位差更小（避免红→绿的对立），
    # 让蓝-黄差异更大（沿安全轴）
    # 相位差自适应：低熵图（对比强烈）用较小相位差保持色温一致，
    # 高熵图（信息丰富）用较大相位差增加色彩区分度
    phase_diff = np.pi * (0.35 + 0.2 * e)  # 范围 [0.35π, 0.55π]
    phase_r = phase_offset + 0.2  # 偏向偏暖
    phase_g = phase_offset + 0.2 + phase_diff  # 与 R 的差由直方图熵决定
    phase_b = phase_offset + np.pi * 1.0  # 与 R/G 相反 → 蓝-黄主导

    return freq, phase_r, phase_g, phase_b


def _confusion_axis_loss(rgb, cvd_type="deutan"):
    """
    创新点 1：计算伪彩色图像在色盲混淆轴上的色差损失。

    色盲患者看不到红-绿轴信息；本文希望伪彩色映射在被色盲模拟后
    仍保留尽量多的可区分信息。计算 (rgb - sim) 的平均范数：
    若大，说明大量信息被损失在混淆轴上；若小，说明映射主要沿安全轴。

    Returns:
        float, 越小越好
    """
    sim = simulate_cvd(rgb, cvd_type)
    diff = np.abs(rgb - sim)
    return float(diff.mean())


def _candidate_palette_search(gray, stats, cvd_type="deutan"):
    """
    创新点 1+2：在 sin-based 候选 + viridis 候选中搜索综合损失最小者。

    综合损失 = max(混淆轴损失_protan, 混淆轴损失_deutan) + λ_nat * MSE(候选, viridis)
    - 同时考虑两种色盲类型，避免对单一类型过拟合
    - 以 viridis 为自然性基准（感知均匀、色盲友好），而非 jet

    Returns:
        (best_rgb_palette, best_method_name, best_loss)
    """
    candidates = []
    viridis_rgb = m6_viridis(gray)

    # 候选 1: 自适应 sin-based
    freq, pr, pg, pb = _adaptive_sin_params(stats)
    sin_rgb = m2_sin_three_phase(gray, freq_r=freq, freq_g=freq, freq_b=freq,
                                  phase_r=pr, phase_g=pg, phase_b=pb)
    candidates.append(("adaptive_sin", sin_rgb))

    # 候选 2: viridis (perceptually uniform, 已经较色盲友好)
    candidates.append(("viridis", viridis_rgb))

    # 候选 3: sin-based 蓝-黄变体（强制 R 和 G 同步，B 反相）
    sin2_rgb = m2_sin_three_phase(gray, freq_r=0.7, freq_g=0.7, freq_b=0.7,
                                   phase_r=0.0, phase_g=0.3, phase_b=np.pi)
    candidates.append(("sin_blue_yellow", sin2_rgb))

    # 评分：综合考虑两种色盲类型的混淆轴损失 + 与 viridis 的自然性约束
    lambda_nat = 0.25
    best = None
    best_loss = float("inf")
    best_name = ""
    for name, rgb in candidates:
        loss_protan = _confusion_axis_loss(rgb, "protan")
        loss_deutan = _confusion_axis_loss(rgb, "deutan")
        cvd_loss = max(loss_protan, loss_deutan)  # 取两种色盲中的最差情况
        nat_loss = float(np.mean((rgb - viridis_rgb) ** 2))  # 与 viridis 的 MSE
        total_loss = (1 - lambda_nat) * cvd_loss + lambda_nat * nat_loss
        if total_loss < best_loss:
            best_loss = total_loss
            best = rgb
            best_name = name

    return best, best_name, best_loss


def m9_proposed(rgb_or_gray, cvd_type="deutan",
                use_axis_avoidance=True,
                use_adaptive=True,
                use_saliency=True,
                saliency_alpha=0.5,
                hsi_lightness_constraint=True):
    """
    Historical fusion baseline used in the final table as M9_old.

    Args:
        rgb_or_gray: 输入灰度图 H×W 或 H×W×3
        cvd_type: 目标色盲类型 ("protan" / "deutan" / "tritan")
        use_axis_avoidance: 是否启用创新点 1（混淆轴避让）
        use_adaptive: 是否启用创新点 2（直方图自适应）
        use_saliency: 是否启用创新点 3（显著性增强）
        saliency_alpha: 显著性增强强度
        hsi_lightness_constraint: 是否约束 V 通道与原灰度一致

    Returns:
        H×W×3 float [0,1]
    """
    arr = np.asarray(rgb_or_gray)
    # 转灰度（始终基于灰度做伪彩色）
    if arr.ndim == 3 and arr.shape[2] == 3:
        if arr.dtype == np.uint8:
            gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0
        else:
            gray_u = (np.clip(arr, 0, 1) * 255).astype(np.uint8)
            gray = cv2.cvtColor(gray_u, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0
    else:
        gray = _to_float01(arr)

    # 1. 直方图统计
    stats = _histogram_statistics(gray)

    # 2. 候选搜索 + 混淆轴最小化
    if use_axis_avoidance and use_adaptive:
        rgb_base, _, _ = _candidate_palette_search(gray, stats, cvd_type)
    elif use_adaptive:
        # 仅自适应 sin-based，不搜索
        freq, pr, pg, pb = _adaptive_sin_params(stats)
        rgb_base = m2_sin_three_phase(gray, freq_r=freq, freq_g=freq, freq_b=freq,
                                       phase_r=pr, phase_g=pg, phase_b=pb)
    else:
        # 退化到固定 viridis
        rgb_base = m6_viridis(gray)

    # 3. 显著性区域增强
    if use_saliency:
        sal = saliency_map(gray, alpha=0.6, win=15)
        # 在显著区域增强对比度（向通道中位数做"反向"拉伸）
        median_per_channel = np.median(rgb_base.reshape(-1, 3), axis=0)
        rgb_enhanced = rgb_base + saliency_alpha * sal[..., None] * (rgb_base - median_per_channel)
        rgb_enhanced = np.clip(rgb_enhanced, 0, 1)
    else:
        rgb_enhanced = rgb_base

    # 4. YCbCr 亮度约束：用真正的亮度通道 Y 替换为原灰度
    #    HSV 的 V = max(R,G,B) 不是感知亮度，Y = 0.299R+0.587G+0.114B 才是
    if hsi_lightness_constraint:
        rgb_clamped = np.clip(rgb_enhanced, 0, 1)
        # RGB → YCbCr
        Y  =  0.299 * rgb_clamped[..., 0] + 0.587 * rgb_clamped[..., 1] + 0.114 * rgb_clamped[..., 2]
        Cb = -0.169 * rgb_clamped[..., 0] - 0.331 * rgb_clamped[..., 1] + 0.500 * rgb_clamped[..., 2] + 0.5
        Cr =  0.500 * rgb_clamped[..., 0] - 0.419 * rgb_clamped[..., 1] - 0.081 * rgb_clamped[..., 2] + 0.5
        # Y → 原灰度
        Y = gray
        # YCbCr → RGB
        r = Y + 1.402 * (Cr - 0.5)
        g = Y - 0.344 * (Cb - 0.5) - 0.714 * (Cr - 0.5)
        b = Y + 1.772 * (Cb - 0.5)
        rgb_final = np.stack([r, g, b], axis=-1)
    else:
        rgb_final = rgb_enhanced

    return np.clip(rgb_final, 0, 1)


if __name__ == "__main__":
    # 自检：在 X-ray synthetic 上测 M9 + 三个消融变体
    import matplotlib.pyplot as plt
    from PIL import Image

    img = np.array(Image.open("../data/grayscale/xray_synthetic.png").convert("L")) / 255.0

    fig, axes = plt.subplots(2, 3, figsize=(12, 8))
    axes[0, 0].imshow(img, cmap="gray"); axes[0, 0].set_title("Original gray"); axes[0, 0].axis("off")

    full = m9_proposed(img, "deutan")
    axes[0, 1].imshow(full); axes[0, 1].set_title("M9 full (proposed)"); axes[0, 1].axis("off")

    sim_full = simulate_cvd(full, "deutan")
    axes[0, 2].imshow(sim_full); axes[0, 2].set_title("M9 seen by deutan"); axes[0, 2].axis("off")

    no_axis = m9_proposed(img, "deutan", use_axis_avoidance=False)
    axes[1, 0].imshow(no_axis); axes[1, 0].set_title("Ablation: no axis-avoidance"); axes[1, 0].axis("off")

    no_adapt = m9_proposed(img, "deutan", use_adaptive=False)
    axes[1, 1].imshow(no_adapt); axes[1, 1].set_title("Ablation: no adaptive"); axes[1, 1].axis("off")

    no_sal = m9_proposed(img, "deutan", use_saliency=False)
    axes[1, 2].imshow(no_sal); axes[1, 2].set_title("Ablation: no saliency"); axes[1, 2].axis("off")

    plt.tight_layout()
    plt.savefig("../figs/_test_proposed.png", dpi=100, bbox_inches="tight")
    print("Saved figs/_test_proposed.png")

    # 量化检查
    from cvd_simulate import simulate_cvd as sim
    print(f"M9 vs original mean color diff: {np.mean(np.abs(full - np.stack([img]*3, axis=-1))):.4f}")
    print(f"M9 axis-loss (deutan): {_confusion_axis_loss(full, 'deutan'):.4f}")
    print(f"M2 sin axis-loss (deutan): {_confusion_axis_loss(m2_sin_three_phase(img), 'deutan'):.4f}")
