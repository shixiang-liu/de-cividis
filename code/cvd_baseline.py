"""
cvd_baseline.py — 色盲适配基线方法

实现：
  M3: HSI 色相旋转   — 课件 Ch6.2 HSI + 鲍吉斌方法（H 通道旋转 120°）
  M7: Brettel + Fidaner Daltonization (1997 / 2005)
  M8: Saliency-Aware Recoloring (InnoColor / WACV 2025 思想轻量复现)

输入彩色 (H×W×3 [0,1])，输出彩色。
对灰度输入，先用一种 baseline pseudocolor 做底，再做色盲适配。
"""
import numpy as np
import cv2

from cvd_simulate import simulate_cvd, srgb_to_linear, linear_to_srgb, RGB2LMS, LMS2RGB
from saliency import saliency_map
from pseudocolor_classic import m4_jet, _to_float01


def _ensure_rgb_float01(rgb_or_gray):
    """灰度 → jet 伪彩色；彩色 → 归一化 float [0,1] H×W×3"""
    arr = np.asarray(rgb_or_gray)
    if arr.ndim == 2 or (arr.ndim == 3 and arr.shape[2] == 1):
        return m4_jet(arr).astype(np.float64)
    arr = arr.astype(np.float64)
    if arr.max() > 1.5:
        arr = arr / 255.0
    return np.clip(arr, 0, 1)


# =================== M3: HSI 色相旋转 ===================
def m3_hsi_hue_rotation(rgb_or_gray, rotation_deg=120):
    """
    课件 Ch6.2 HSI 色彩空间 + 鲍吉斌方法（H 通道旋转）。

    若输入是灰度，先用 jet 生成基础彩图再旋转色相；
    若输入是彩色，直接旋转 H 通道。

    Args:
        rgb_or_gray: H×W (灰度) 或 H×W×3 (彩色)
        rotation_deg: 色相旋转角度（默认 120°）

    Returns:
        H×W×3 float [0,1]
    """
    rgb = _ensure_rgb_float01(rgb_or_gray).astype(np.float32)
    hsv = cv2.cvtColor((rgb * 255).astype(np.uint8), cv2.COLOR_RGB2HSV).astype(np.float32)
    # OpenCV 的 H 取值在 [0, 180]，对应 0-360°
    hsv[..., 0] = (hsv[..., 0] + rotation_deg / 2.0) % 180
    rgb_out = cv2.cvtColor(hsv.astype(np.uint8), cv2.COLOR_HSV2RGB).astype(np.float64) / 255.0
    return np.clip(rgb_out, 0, 1)


# =================== M7: Brettel + Fidaner Daltonization ===================
# 误差矩阵：把红绿混淆轴上的色差投影到蓝-黄安全轴
# 原始 Fidaner 2005 误差扩散矩阵（针对 Deutan）
ERROR_MAT_DEUTAN = np.array([
    [0,    0,    0],
    [0.7,  1,    0],
    [0.7,  0,    1],
], dtype=np.float64)

ERROR_MAT_PROTAN = np.array([
    [0,    0,    0],
    [0.7,  1,    0],
    [0.7,  0,    1],
], dtype=np.float64)

ERROR_MAT_TRITAN = np.array([
    [1,    0,    0.7],
    [0,    1,    0.7],
    [0,    0,    0],
], dtype=np.float64)

ERR_MATS = {"protan": ERROR_MAT_PROTAN, "deutan": ERROR_MAT_DEUTAN, "tritan": ERROR_MAT_TRITAN}


def m7_daltonize(rgb_or_gray, cvd_type="deutan"):
    """
    Brettel-Mollon 模拟 + Fidaner 2005 Daltonization 误差扩散。

    Args:
        rgb_or_gray: H×W or H×W×3
        cvd_type: "protan" / "deutan" / "tritan"

    Returns:
        H×W×3 float [0,1] —— 经过色盲补偿的图像
    """
    rgb = _ensure_rgb_float01(rgb_or_gray)

    # 1. 模拟色盲所见
    sim = simulate_cvd(rgb, cvd_type)

    # 2. 误差 = 原图 - 模拟图（这部分信息色盲患者看不到）
    err = rgb - sim

    # 3. 把误差通过误差矩阵投影到安全色轴（蓝-黄）
    err_redistributed = err.reshape(-1, 3) @ ERR_MATS[cvd_type].T
    err_redistributed = err_redistributed.reshape(rgb.shape)

    # 4. 加回原图
    out = rgb + err_redistributed
    return np.clip(out, 0, 1)


# =================== M8: Saliency-Aware Recoloring ===================
def m8_saliency_aware_recolor(rgb_or_gray, cvd_type="deutan", saliency_weight=1.5):
    """
    InnoColor / WACV 2025 思想的轻量复现：在显著性区域优先做色盲补偿，
    在非显著区域保持原色彩。

    Args:
        rgb_or_gray: 输入图像
        cvd_type: 色盲类型
        saliency_weight: 显著区域的补偿强度倍数

    Returns:
        H×W×3 float [0,1]
    """
    arr = np.asarray(rgb_or_gray)
    rgb = _ensure_rgb_float01(arr)
    # 用亮度做显著性
    gray_for_sal = (0.299 * rgb[..., 0] + 0.587 * rgb[..., 1] + 0.114 * rgb[..., 2])
    gray_for_sal = (gray_for_sal * 255).astype(np.uint8)

    # 计算显著性图
    sal = saliency_map(gray_for_sal, alpha=0.6, win=15)  # H×W float [0,1]

    # 全图 Daltonization
    daltonized = m7_daltonize(rgb, cvd_type)

    # 显著性加权融合
    weight = np.clip(sal * saliency_weight, 0, 1)[..., None]
    out = weight * daltonized + (1 - weight) * rgb
    return np.clip(out, 0, 1)


CVD_BASELINES = {
    "M3_hsi_rotation":   m3_hsi_hue_rotation,
    "M7_daltonize":      m7_daltonize,
    "M8_saliency_aware": m8_saliency_aware_recolor,
}


if __name__ == "__main__":
    import matplotlib.pyplot as plt
    from PIL import Image

    img = np.array(Image.open("../data/usc-sipi/lena.png").convert("RGB")) / 255.0
    fig, axes = plt.subplots(1, 4, figsize=(14, 4))
    axes[0].imshow(img); axes[0].set_title("Original Lena"); axes[0].axis("off")
    for ax, (name, fn) in zip(axes[1:], CVD_BASELINES.items()):
        out = fn(img, "deutan") if name != "M3_hsi_rotation" else fn(img)
        ax.imshow(out); ax.set_title(name); ax.axis("off")
    plt.tight_layout()
    plt.savefig("../figs/_test_cvd_baseline.png", dpi=100, bbox_inches="tight")
    print("Saved figs/_test_cvd_baseline.png")
