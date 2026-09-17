"""
metrics.py — 实验评价指标

五个指标：
  1) ΔE2000 (CIEDE2000): 感知色差，评估自然性
  2) Y-SSIM: 亮度通道结构相似性
  3) RGB-SSIM: RGB 三通道结构相似性
  4) AvgGrad: 色觉障碍模拟后的平均梯度强度
  5) 信息熵 H: 彩色映射后图像的灰度信息熵

输入约定：
  rgb1, rgb2: H×W×3 float [0,1] sRGB 图像
  返回 float 标量
"""
import numpy as np
from skimage.color import rgb2lab, deltaE_ciede2000
from skimage.metrics import structural_similarity as ssim_skimage

from cvd_simulate import simulate_cvd
from saliency import sobel_gradient_magnitude


def _ensure_float01(rgb):
    arr = np.asarray(rgb, dtype=np.float64)
    if arr.max() > 1.5:
        arr = arr / 255.0
    return np.clip(arr, 0, 1)


def delta_e2000(rgb1, rgb2):
    """CIEDE2000 mean color difference (越小越保形)"""
    a = _ensure_float01(rgb1)
    b = _ensure_float01(rgb2)
    # 若一方是灰度图，转成 RGB 再算
    if a.ndim == 2:
        a = np.stack([a, a, a], axis=-1)
    if b.ndim == 2:
        b = np.stack([b, b, b], axis=-1)
    lab1 = rgb2lab(a)
    lab2 = rgb2lab(b)
    de = deltaE_ciede2000(lab1, lab2)
    return float(np.mean(de))


def ssim(rgb1, rgb2):
    """亮度通道 SSIM (越大越相似, 1.0 完全一致)。

    将两幅图转到 YCbCr 空间后只比较 Y（亮度）通道，
    衡量结构保真度，不受颜色差异影响。
    """
    a = _ensure_float01(rgb1)
    b = _ensure_float01(rgb2)
    if a.ndim == 3:
        ya = 0.299 * a[..., 0] + 0.587 * a[..., 1] + 0.114 * a[..., 2]
    else:
        ya = a
    if b.ndim == 3:
        yb = 0.299 * b[..., 0] + 0.587 * b[..., 1] + 0.114 * b[..., 2]
    else:
        yb = b
    return float(ssim_skimage(ya, yb, data_range=1.0))


def ssim_rgb(rgb1, rgb2):
    """RGB 三通道分别计算 SSIM 取均值 (越大越相似)。

    论文中 RGB-SSIM 的计算方式：对 R/G/B 三个通道分别算 SSIM 再平均。
    """
    a = _ensure_float01(rgb1)
    b = _ensure_float01(rgb2)
    if a.ndim == 2:
        a = np.stack([a, a, a], axis=-1)
    if b.ndim == 2:
        b = np.stack([b, b, b], axis=-1)
    vals = []
    for c in range(3):
        vals.append(float(ssim_skimage(a[..., c], b[..., c], data_range=1.0)))
    return float(np.mean(vals))


def average_gradient(rgb_mapped, cvd_type="deutan"):
    """色觉障碍模拟后的平均 Sobel 梯度强度。

    该指标只描述边缘和纹理强度，数值越大表示局部变化越明显；
    它不能单独代表视觉质量，因此需要与色差、SSIM 和图像对比一起看。
    """
    sim_mapped = simulate_cvd(_ensure_float01(rgb_mapped), cvd_type)
    gray_m = (0.299 * sim_mapped[..., 0] + 0.587 * sim_mapped[..., 1] +
              0.114 * sim_mapped[..., 2]).astype(np.float32)
    return float(sobel_gradient_magnitude(gray_m).mean())


def color_entropy(rgb):
    """彩色图像信息熵 (3 通道熵之和) — 越大表示色彩信息越丰富。"""
    arr = _ensure_float01(rgb)
    total = 0.0
    for c in range(3):
        ch = (arr[..., c] * 255).astype(np.uint8)
        hist, _ = np.histogram(ch, bins=256, range=(0, 256))
        p = hist.astype(np.float64) / max(hist.sum(), 1)
        p_nz = p[p > 0]
        h = -(p_nz * np.log2(p_nz)).sum()
        total += h
    return float(total)


def all_metrics(rgb_mapped, gray_ref=None, cvd_type="deutan"):
    """
    一次性计算所有指标。

    Args:
        rgb_mapped: 待评估的伪彩色图
        gray_ref: 原灰度图 (H×W float [0,1])，用于 ΔE2000 和 SSIM
        cvd_type: 色盲类型

    Returns:
        dict with keys: delta_e2000, ssim, ssim_rgb, avg_gradient, entropy
    """
    if gray_ref is not None:
        de = delta_e2000(rgb_mapped, gray_ref)
        s  = ssim(rgb_mapped, gray_ref)
        s_rgb = ssim_rgb(rgb_mapped, gray_ref)
    else:
        de = 0.0
        s  = 1.0
        s_rgb = 1.0

    ag = average_gradient(rgb_mapped, cvd_type)
    he = color_entropy(rgb_mapped)
    return {"delta_e2000": de, "ssim": s, "ssim_rgb": s_rgb, "avg_gradient": ag, "entropy": he}


if __name__ == "__main__":
    # 自检
    from PIL import Image
    from pseudocolor_classic import m2_sin_three_phase, m4_jet, m6_viridis
    from pseudocolor_proposed import m9_proposed

    img = np.array(Image.open("../data/grayscale/xray_synthetic.png").convert("L")) / 255.0

    out_jet = m4_jet(img)
    out_viridis = m6_viridis(img)
    out_m9 = m9_proposed(img, "deutan")

    print("=== 各方法对比 (deutan 视角) ===")
    print("方法           | ΔE vs jet | SSIM vs jet | AvgGrad | Entropy")
    for name, out in [("jet (M4)", out_jet), ("viridis (M6)", out_viridis), ("M9 proposed", out_m9)]:
        m = all_metrics(out, out_jet, "deutan")
        print(f"{name:14s} | {m['delta_e2000']:9.2f} | {m['ssim']:11.3f} | "
              f"{m['avg_gradient']:8.3f} | {m['entropy']:.2f}")
