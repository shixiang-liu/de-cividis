"""
cvd_simulate.py — 色盲视觉模拟 (Brettel-Mollon-Mollon 1997)

将正常视觉图像映射为色盲患者所见图像。基于 LMS 色彩空间和混淆轴投影。

参考文献：
  Brettel H, Viénot F, Mollon J D. Computerized simulation of color appearance
  for dichromats. JOSA A, 1997, 14(10): 2647-2655.

使用：
    from cvd_simulate import simulate_cvd
    sim_img = simulate_cvd(rgb_img, cvd_type='deutan')  # rgb_img: float in [0,1]
"""
import numpy as np

# sRGB → LMS（Hunt-Pointer-Estevez D65 调整后矩阵）
RGB2LMS = np.array([
    [17.8824,   43.5161,   4.1194],
    [3.45565,   27.1554,   3.86714],
    [0.0299566, 0.184309, 1.46709],
], dtype=np.float64)

LMS2RGB = np.linalg.inv(RGB2LMS)

# Brettel-Mollon-Mollon 1997 色盲混淆面投影矩阵
# Protanopia (red-blind, L cone missing)
PROTAN_SIM = np.array([
    [0.0,    2.02344, -2.52581],
    [0.0,    1.0,      0.0],
    [0.0,    0.0,      1.0],
], dtype=np.float64)

# Deuteranopia (green-blind, M cone missing)
DEUTAN_SIM = np.array([
    [1.0,      0.0,    0.0],
    [0.494207, 0.0,    1.24827],
    [0.0,      0.0,    1.0],
], dtype=np.float64)

# Tritanopia (blue-blind, S cone missing)
TRITAN_SIM = np.array([
    [1.0,       0.0,       0.0],
    [0.0,       1.0,       0.0],
    [-0.395913, 0.801109,  0.0],
], dtype=np.float64)

CVD_MATRICES = {
    "protan":  PROTAN_SIM,
    "deutan":  DEUTAN_SIM,
    "tritan":  TRITAN_SIM,
}


def srgb_to_linear(srgb):
    """sRGB gamma decoding to linear RGB. Input/output in [0,1]."""
    a = 0.055
    return np.where(srgb <= 0.04045, srgb / 12.92, ((srgb + a) / (1 + a)) ** 2.4)


def linear_to_srgb(lin):
    """Linear RGB to sRGB encoding. Input/output in [0,1]."""
    a = 0.055
    return np.where(lin <= 0.0031308, lin * 12.92, (1 + a) * lin ** (1 / 2.4) - a)


def simulate_cvd(rgb_img, cvd_type="deutan"):
    """
    模拟色盲视觉。

    Args:
        rgb_img: H×W×3 numpy array, sRGB values in [0,1]
        cvd_type: "protan" / "deutan" / "tritan"

    Returns:
        H×W×3 sRGB image showing what a CVD patient sees
    """
    if cvd_type not in CVD_MATRICES:
        raise ValueError(f"cvd_type must be one of {list(CVD_MATRICES)}")

    img = np.asarray(rgb_img, dtype=np.float64)
    if img.ndim == 2:
        img = np.stack([img, img, img], axis=-1)
    if img.max() > 1.5:
        img = img / 255.0

    # sRGB → linear RGB
    lin = srgb_to_linear(np.clip(img, 0, 1))

    # linear RGB → LMS
    lms = lin.reshape(-1, 3) @ RGB2LMS.T

    # Project onto CVD confusion plane
    sim_lms = lms @ CVD_MATRICES[cvd_type].T

    # LMS → linear RGB
    sim_lin = sim_lms @ LMS2RGB.T
    sim_lin = sim_lin.reshape(img.shape)
    sim_lin = np.clip(sim_lin, 0, 1)

    # linear RGB → sRGB
    sim_srgb = linear_to_srgb(sim_lin)
    return np.clip(sim_srgb, 0, 1)


def cvd_simulate_uint8(rgb_img_u8, cvd_type="deutan"):
    """便捷接口: 接收 uint8 图像, 返回 uint8 图像."""
    img = rgb_img_u8.astype(np.float64) / 255.0
    sim = simulate_cvd(img, cvd_type)
    return (sim * 255).astype(np.uint8)


if __name__ == "__main__":
    # 自检：生成红绿蓝色块测试图
    import matplotlib.pyplot as plt
    test = np.zeros((100, 300, 3))
    test[:, :100, 0] = 1.0    # red
    test[:, 100:200, 1] = 1.0 # green
    test[:, 200:, 2] = 1.0    # blue
    fig, ax = plt.subplots(1, 4, figsize=(12, 3))
    ax[0].imshow(test); ax[0].set_title("Original"); ax[0].axis("off")
    for i, t in enumerate(["protan", "deutan", "tritan"], 1):
        sim = simulate_cvd(test, t)
        ax[i].imshow(sim); ax[i].set_title(t); ax[i].axis("off")
    plt.tight_layout()
    plt.savefig("../figs/_test_cvd_simulate.png", dpi=100, bbox_inches="tight")
    print("Saved figs/_test_cvd_simulate.png")
