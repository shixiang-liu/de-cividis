"""
pseudocolor_classic.py — 经典伪彩色映射方法

实现：
  M1: 密度分层 (Intensity Slicing) — 课件 Ch6.3 p37
  M2: sin-based 三相位变换       — 课件 Ch6.3 p44-46
  M4: Jet (Rainbow) Colormap       — Matplotlib 经典
  M5: Hot Colormap                 — Matplotlib 经典
  M6: Viridis Colormap             — Matplotlib 2.0+ perceptually uniform

所有方法接收单通道灰度图 (H×W, float [0,1] 或 uint8)，
返回彩色图像 (H×W×3, float [0,1])。
"""
import numpy as np
import matplotlib.cm as cm


def _to_float01(gray):
    """转换为 float [0,1] 单通道。"""
    arr = np.asarray(gray)
    if arr.ndim == 3:
        # RGB 转灰度
        arr = 0.299 * arr[..., 0] + 0.587 * arr[..., 1] + 0.114 * arr[..., 2]
    if arr.dtype != np.float64 and arr.dtype != np.float32:
        arr = arr.astype(np.float64) / 255.0
    return np.clip(arr.astype(np.float64), 0, 1)


# ===== M1: 密度分层 (Intensity Slicing) =====
def m1_intensity_slicing(gray, n_levels=8, palette=None):
    """
    课件 Ch6.3 p37 密度分层：把灰度区间划分为 n 段离散色块。

    Args:
        gray: H×W 灰度图
        n_levels: 分层数（默认 8）
        palette: H×W×3 调色板，shape=(n_levels, 3) float [0,1]

    Returns:
        H×W×3 float [0,1]
    """
    g = _to_float01(gray)
    if palette is None:
        # 默认采用课件示意的 8 段彩色调色板
        palette = np.array([
            [0.00, 0.00, 0.50],   # 0: 深蓝
            [0.00, 0.40, 1.00],   # 1: 蓝
            [0.00, 1.00, 1.00],   # 2: 青
            [0.00, 1.00, 0.30],   # 3: 黄绿
            [1.00, 1.00, 0.00],   # 4: 黄
            [1.00, 0.55, 0.00],   # 5: 橙
            [1.00, 0.00, 0.00],   # 6: 红
            [0.55, 0.00, 0.55],   # 7: 紫红
        ], dtype=np.float64)
    palette = np.asarray(palette, dtype=np.float64)
    n_levels = palette.shape[0]
    idx = np.clip((g * n_levels).astype(int), 0, n_levels - 1)
    out = palette[idx]
    return out


# ===== M2: sin-based 三相位变换 (课件 Ch6.3 p44-46) =====
def m2_sin_three_phase(gray, freq_r=1.0, freq_g=1.0, freq_b=1.0,
                      phase_r=0.0, phase_g=2.0944, phase_b=4.1888):
    """
    课件 Ch6.3 p44-46 描述的"利用各正弦型的相位和频率变化做彩色变换"。

    R(I) = 0.5 + 0.5 * sin(2π * f_r * I + φ_r)
    G(I) = 0.5 + 0.5 * sin(2π * f_g * I + φ_g)
    B(I) = 0.5 + 0.5 * sin(2π * f_b * I + φ_b)

    默认相位：0, 2π/3, 4π/3（即 0°, 120°, 240°）

    Args:
        gray: H×W 灰度
        freq_r/g/b: 三通道频率（同频不同相位 → 单色到彩色映射）
        phase_r/g/b: 三通道相位

    Returns:
        H×W×3 float [0,1]
    """
    g = _to_float01(gray)
    R = 0.5 + 0.5 * np.sin(2 * np.pi * freq_r * g + phase_r)
    G = 0.5 + 0.5 * np.sin(2 * np.pi * freq_g * g + phase_g)
    B = 0.5 + 0.5 * np.sin(2 * np.pi * freq_b * g + phase_b)
    out = np.stack([R, G, B], axis=-1)
    return np.clip(out, 0, 1)


# ===== M4: Jet (Rainbow) =====
def m4_jet(gray):
    """Matplotlib jet colormap（经典彩虹色 — 色盲不友好的对照基线）。"""
    g = _to_float01(gray)
    return cm.jet(g)[..., :3]  # 去掉 alpha


# ===== M5: Hot =====
def m5_hot(gray):
    """Matplotlib hot colormap（黑→红→黄→白）。"""
    g = _to_float01(gray)
    return cm.hot(g)[..., :3]


def m5b_turbo(gray):
    """Google/Matplotlib Turbo colormap: smoother rainbow-style baseline."""
    g = _to_float01(gray)
    return cm.turbo(g)[..., :3]


# ===== M6: Viridis =====
def m6_viridis(gray):
    """Matplotlib viridis colormap（现代 perceptually uniform）。"""
    g = _to_float01(gray)
    return cm.viridis(g)[..., :3]


def m6_cividis(gray):
    """Matplotlib cividis colormap: CVD-friendly sequential baseline."""
    g = _to_float01(gray)
    return cm.cividis(g)[..., :3]


# ===== 简易统一接口 =====
CLASSIC_METHODS = {
    "M1_intensity_slicing": m1_intensity_slicing,
    "M2_sin_three_phase":   m2_sin_three_phase,
    "M4_jet":               m4_jet,
    "M5_hot":               m5_hot,
    "M5b_turbo":            m5b_turbo,
    "M6_viridis":           m6_viridis,
    "M6b_cividis":          m6_cividis,
}


if __name__ == "__main__":
    # 自检：在合成灰度梯度图上展示 5 种方法
    import matplotlib.pyplot as plt

    # 灰度梯度图
    grad = np.linspace(0, 1, 512).reshape(1, -1).repeat(80, axis=0)

    fig, axes = plt.subplots(len(CLASSIC_METHODS) + 1, 1, figsize=(10, 10))
    axes[0].imshow(grad, cmap="gray", aspect="auto"); axes[0].set_title("原灰度梯度"); axes[0].axis("off")
    for ax, (name, fn) in zip(axes[1:], CLASSIC_METHODS.items()):
        out = fn(grad)
        ax.imshow(out, aspect="auto"); ax.set_title(name); ax.axis("off")
    plt.tight_layout()
    plt.savefig("../figs/_test_classic_pseudocolor.png", dpi=100, bbox_inches="tight")
    print("Saved figs/_test_classic_pseudocolor.png")
