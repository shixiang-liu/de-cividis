"""
saliency.py — 显著性图构建（融合课件 Ch3.7 Sobel + Ch3.3.4 局部统计量）

显著性 S(x,y) = α · |∇I|_norm + (1-α) · LocalContrast(I)

其中：
- |∇I| 用 Sobel 算子（课件 Ch3.7.3 公式 3.7-1 / 3.7-2）
- LocalContrast 用局部均值与全局均值的偏离 + 局部标准差（课件 Ch3.3.4 思想）

输出归一化到 [0, 1]，作为 M9 本文方法的 saliency map。
"""
import numpy as np
import cv2


def sobel_gradient_magnitude(gray):
    """
    课件 Ch3.7.3 公式 3.7-5(Sobel)：|G| = sqrt(Gx^2 + Gy^2)

    Args:
        gray: H×W float [0,1] or uint8

    Returns:
        H×W float [0,1] gradient magnitude (max-normalized)
    """
    if gray.dtype == np.uint8:
        gray = gray.astype(np.float32) / 255.0
    else:
        gray = gray.astype(np.float32)
    gx = cv2.Sobel(gray, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray, cv2.CV_32F, 0, 1, ksize=3)
    mag = np.sqrt(gx * gx + gy * gy)
    if mag.max() > 0:
        mag = mag / mag.max()
    return mag.astype(np.float64)


def local_statistics(gray, win=15):
    """
    课件 Ch3.3.4 局部均值 m_S(x,y) 与局部方差 D_S(x,y)。

    Args:
        gray: H×W float [0,1] or uint8
        win: 局部窗口大小

    Returns:
        (m_local, std_local), 都是 H×W float
    """
    if gray.dtype == np.uint8:
        gray = gray.astype(np.float32) / 255.0
    else:
        gray = gray.astype(np.float32)
    kernel = (win, win)
    m = cv2.boxFilter(gray, cv2.CV_32F, kernel)
    sq = cv2.boxFilter(gray * gray, cv2.CV_32F, kernel)
    var = np.maximum(sq - m * m, 0.0)
    std = np.sqrt(var)
    return m.astype(np.float64), std.astype(np.float64)


def local_contrast_score(gray, win=15):
    """
    局部对比度评分：基于局部标准差归一化（局部纹理丰富度）+
    局部均值偏离全局均值（局部对比度）。

    Returns:
        H×W float [0,1]
    """
    m_loc, std_loc = local_statistics(gray, win)
    if gray.dtype == np.uint8:
        gray_f = gray.astype(np.float32) / 255.0
    else:
        gray_f = gray.astype(np.float32)
    m_global = float(gray_f.mean())

    # 归一化局部标准差（纹理强度）
    std_norm = std_loc
    if std_norm.max() > 0:
        std_norm = std_norm / std_norm.max()

    # 归一化局部均值与全局均值的偏离
    dev = np.abs(m_loc - m_global)
    if dev.max() > 0:
        dev = dev / dev.max()

    # 加权融合
    score = 0.6 * std_norm + 0.4 * dev
    score = np.clip(score, 0, 1)
    return score


def saliency_map(gray, alpha=0.6, win=15):
    """
    本文显著性图：S = α · |∇I|_norm + (1-α) · LocalContrast(I)

    Args:
        gray: H×W gray image (float [0,1] or uint8)
        alpha: 梯度权重 (默认 0.6)
        win: 局部窗口

    Returns:
        H×W float [0,1] 显著性图
    """
    if gray.ndim == 3:
        # 转灰度（Y = 0.299R + 0.587G + 0.114B，课件 Ch6.5）
        if gray.dtype == np.uint8:
            gray = cv2.cvtColor(gray, cv2.COLOR_RGB2GRAY)
        else:
            gray_u = (np.clip(gray, 0, 1) * 255).astype(np.uint8)
            gray = cv2.cvtColor(gray_u, cv2.COLOR_RGB2GRAY)

    grad = sobel_gradient_magnitude(gray)
    cont = local_contrast_score(gray, win)

    sal = alpha * grad + (1 - alpha) * cont
    sal = np.clip(sal, 0, 1)
    return sal


if __name__ == "__main__":
    # 自检：在 Lena 上可视化显著性图
    import matplotlib.pyplot as plt
    from PIL import Image

    img = np.array(Image.open("../data/usc-sipi/lena.png").convert("L"))
    sal = saliency_map(img, alpha=0.6, win=15)

    fig, ax = plt.subplots(1, 3, figsize=(12, 4))
    ax[0].imshow(img, cmap="gray"); ax[0].set_title("Original (Lena gray)"); ax[0].axis("off")
    ax[1].imshow(sobel_gradient_magnitude(img), cmap="hot"); ax[1].set_title("|∇I| (Sobel)"); ax[1].axis("off")
    ax[2].imshow(sal, cmap="hot"); ax[2].set_title("Saliency S = 0.6·|∇| + 0.4·LocalContrast"); ax[2].axis("off")
    plt.tight_layout()
    plt.savefig("../figs/_test_saliency.png", dpi=100, bbox_inches="tight")
    print("Saved figs/_test_saliency.png")
    print(f"Saliency stats: min={sal.min():.3f} max={sal.max():.3f} mean={sal.mean():.3f}")
