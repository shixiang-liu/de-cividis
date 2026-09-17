"""Generate figures for the DE-Cividis course paper.

The main comparison figures use only classic pseudocolor methods, public
colormaps, and DE-Cividis as the proposed method. Candidate variants such as
CLAHE-Cividis are kept out of the paper-facing figures.
"""

from __future__ import annotations

import csv
from pathlib import Path

import cv2
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

from cvd_simulate import simulate_cvd
from pseudocolor_classic import (
    m1_intensity_slicing,
    m2_sin_three_phase,
    m4_jet,
    m5_hot,
    m5b_turbo,
    m6_viridis,
    m6_cividis,
)
from pseudocolor_proposed import (
    DE_LUMINANCE_MIX,
    de_cividis,
    de_cividis_no_de,
    de_cividis_no_lw,
    de_jet_no_cividis,
    m16_unsharp_cividis,
)
from run_final_experiments import collect_images


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
FIGS = ROOT / "figs"
RESULTS = ROOT / "results"

SAVE_DPI = 450

mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "SimSun", "DejaVu Sans"]
mpl.rcParams["font.family"] = "sans-serif"
mpl.rcParams["axes.unicode_minus"] = False


def load_gray(path: Path, max_dim: int = 640) -> np.ndarray:
    img = np.array(Image.open(path).convert("RGB"))
    h, w = img.shape[:2]
    if max(h, w) > max_dim:
        scale = max_dim / max(h, w)
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    return cv2.cvtColor(img, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0


def gray_to_rgb(gray: np.ndarray) -> np.ndarray:
    g = np.clip(gray, 0.0, 1.0)
    return np.stack([g, g, g], axis=-1)


def normalize01(arr: np.ndarray) -> np.ndarray:
    x = np.asarray(arr, dtype=np.float64)
    lo = float(x.min())
    hi = float(x.max())
    if hi <= lo + 1e-12:
        return np.zeros_like(x)
    return (x - lo) / (hi - lo)


def save_figure(name: str) -> None:
    FIGS.mkdir(exist_ok=True)
    out = FIGS / name
    plt.savefig(out, dpi=SAVE_DPI, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"saved {out.relative_to(ROOT)}")


def add_subtitle(ax, text: str, size: float = 9.8, weight: str | None = None) -> None:
    ax.set_title(text, fontsize=size, weight=weight, pad=5)


def imshow(ax, image: np.ndarray, title: str | None = None) -> None:
    ax.imshow(np.clip(image, 0.0, 1.0), interpolation="nearest")
    ax.axis("off")
    if title:
        add_subtitle(ax, title)


def fig1_pipeline() -> None:
    """Figure 1: compact method flow."""
    sample = load_gray(DATA / "grayscale" / "xray_synthetic.png", max_dim=260)
    blur = cv2.GaussianBlur(sample, (0, 0), sigmaX=2.0, sigmaY=2.0)
    detail = sample - blur
    enhanced = np.clip(sample + 2.0 * detail, 0.0, 1.0)
    cividis = m6_cividis(enhanced)
    final = de_cividis(sample)

    images = [
        ("输入 I", gray_to_rgb(sample)),
        ("背景层 B", gray_to_rgb(blur)),
        ("细节层 D", gray_to_rgb(normalize01(detail) * 0.8 + 0.1)),
        ("增强灰度 I'", gray_to_rgb(enhanced)),
        ("Cividis 映射", cividis),
        (f"亮度回写 λ={DE_LUMINANCE_MIX:.1f}", final),
    ]

    fig, axes = plt.subplots(1, len(images), figsize=(12.2, 2.25))
    for ax, (title, img) in zip(axes, images):
        imshow(ax, img, title)

    fig.text(0.50, 0.05, "细节增强 -> Cividis 映射 -> YCbCr 亮度回写", ha="center", fontsize=10.5)
    plt.subplots_adjust(left=0.03, right=0.98, top=0.80, bottom=0.25, wspace=0.22)
    for i in range(len(images) - 1):
        left = axes[i].get_position()
        right = axes[i + 1].get_position()
        x = (left.x1 + right.x0) / 2
        y = (left.y0 + left.y1) / 2
        fig.text(x, y, "->", ha="center", va="center", fontsize=14, color="#475569")
    save_figure("fig1_pipeline.png")


def fig2_colormap_cvd() -> None:
    """Figure 2: compare public colormaps under red-green CVD simulation."""
    gray = load_gray(DATA / "grayscale" / "xray_synthetic.png", max_dim=420)
    methods = [
        ("Jet", m4_jet(gray)),
        ("Turbo", m5b_turbo(gray)),
        ("Viridis", m6_viridis(gray)),
        ("Cividis", m6_cividis(gray)),
    ]
    views = [
        ("正常视觉", lambda rgb: rgb),
        ("protan 模拟", lambda rgb: simulate_cvd(rgb, "protan")),
        ("deutan 模拟", lambda rgb: simulate_cvd(rgb, "deutan")),
    ]

    fig, axes = plt.subplots(len(methods), len(views) + 1, figsize=(10.8, 7.4))
    for r, (method_name, rgb) in enumerate(methods):
        imshow(axes[r, 0], gray_to_rgb(gray))
        if r == 0:
            add_subtitle(axes[r, 0], "原灰度")
        for c, (view_name, fn) in enumerate(views, start=1):
            imshow(axes[r, c], fn(rgb))
            if r == 0:
                add_subtitle(axes[r, c], view_name)
        fig.text(0.10, 0.82 - r * 0.22, method_name, ha="right", va="center",
                 fontsize=11, weight="bold")
    plt.subplots_adjust(left=0.13, right=0.99, top=0.93, bottom=0.04, wspace=0.04, hspace=0.08)
    save_figure("fig2_colormap_cvd.png")


def fig3_parameter_effect() -> None:
    """Figure 3: show the parameter choice for sigma, alpha, and lambda."""
    gray = load_gray(DATA / "grayscale" / "xray_synthetic.png", max_dim=420)
    sigma_columns = [
        ("原灰度", gray_to_rgb(gray)),
        ("σ=1.0, α=2.0", m16_unsharp_cividis(gray, sigma=1.0, amount=2.0, luminance_mix=DE_LUMINANCE_MIX)),
        ("σ=2.0, α=2.0", m16_unsharp_cividis(gray, sigma=2.0, amount=2.0, luminance_mix=DE_LUMINANCE_MIX)),
        ("σ=3.0, α=2.0", m16_unsharp_cividis(gray, sigma=3.0, amount=2.0, luminance_mix=DE_LUMINANCE_MIX)),
        ("σ=2.0, α=3.0", m16_unsharp_cividis(gray, sigma=2.0, amount=3.0, luminance_mix=DE_LUMINANCE_MIX)),
    ]
    lambda_columns = [
        ("原灰度", gray_to_rgb(gray)),
        ("λ=1.0", m16_unsharp_cividis(gray, sigma=2.0, amount=2.0, luminance_mix=1.0)),
        ("λ=0.9", m16_unsharp_cividis(gray, sigma=2.0, amount=2.0, luminance_mix=0.9)),
        ("λ=0.7", m16_unsharp_cividis(gray, sigma=2.0, amount=2.0, luminance_mix=0.7)),
        ("λ=0.5", m16_unsharp_cividis(gray, sigma=2.0, amount=2.0, luminance_mix=0.5)),
    ]
    rows = [
        ("σ/α 正常视觉", sigma_columns, lambda rgb: rgb),
        ("σ/α deutan", sigma_columns, lambda rgb: simulate_cvd(rgb, "deutan")),
        ("λ 正常视觉", lambda_columns, lambda rgb: rgb),
        ("λ deutan", lambda_columns, lambda rgb: simulate_cvd(rgb, "deutan")),
    ]

    fig, axes = plt.subplots(len(rows), len(sigma_columns), figsize=(12.2, 8.0))
    for r, (row_name, columns, view_fn) in enumerate(rows):
        for c, (title, img) in enumerate(columns):
            imshow(axes[r, c], view_fn(img))
            if r in (0, 2):
                target = ("2.0, α=2.0" in title) or (title == "λ=0.9")
                add_subtitle(axes[r, c], title, size=9.2, weight="bold" if target else None)
        y = 0.82 - r * 0.215
        fig.text(0.105, y, row_name, ha="right", va="center", fontsize=9.6, weight="bold")
    plt.subplots_adjust(left=0.13, right=0.995, top=0.93, bottom=0.035, wspace=0.04, hspace=0.30)
    save_figure("fig3_saliency.png")


def fig4_evaluation_overview() -> None:
    """Figure 4: overview of all 27 images used in the final evaluation."""
    items = [(category, name, gray) for dataset, category, name, gray in collect_images() if dataset == "primary27"]
    items = sorted(items, key=lambda x: (x[0], x[1]))
    label_map = {
        "ct_synthetic": "CT",
        "lena_gray": "Lena",
        "mri_synthetic": "MRI",
        "rainfall_synthetic": "降雨",
        "satellite_synthetic": "遥感",
        "sem_synthetic": "显微",
        "thermal_synthetic": "热成像",
        "weld_synthetic": "焊缝",
        "xray_synthetic": "X 射线",
        "global": "世界地图",
        "shanghai": "线路地图",
        "airplane": "航空图",
        "lena": "Lena",
        "mandrill": "纹理图",
        "peppers": "蔬菜图",
    }
    ncols = 9
    nrows = int(np.ceil(len(items) / ncols))
    fig, axes = plt.subplots(nrows, ncols, figsize=(13.2, 4.6))
    axes = np.asarray(axes).reshape(nrows, ncols)
    for ax in axes.flat:
        ax.axis("off")
    for i, (category, name, gray) in enumerate(items):
        ax = axes.flat[i]
        imshow(ax, de_cividis(gray))
        title = label_map.get(name, name.replace("plate", "色觉板").replace("_synthetic", ""))
        ax.set_title(title, fontsize=6.8, pad=2)
    plt.subplots_adjust(left=0.02, right=0.995, top=0.94, bottom=0.04, wspace=0.025, hspace=0.16)
    save_figure("fig4_eval_overview.png")


def fig5_typical_cases() -> None:
    """Figure 5: selected visible comparisons using public baselines."""
    cases = [
        (DATA / "grayscale" / "xray_synthetic.png", "X 射线"),
        (DATA / "grayscale" / "weld_synthetic.png", "焊缝"),
        (DATA / "usc-sipi" / "mandrill.png", "纹理"),
    ]
    columns = [
        ("原灰度", gray_to_rgb),
        ("Jet", m4_jet),
        ("Turbo", m5b_turbo),
        ("Viridis", m6_viridis),
        ("Cividis", m6_cividis),
        ("DE-Cividis", de_cividis),
    ]

    fig, axes = plt.subplots(len(cases), len(columns), figsize=(12.8, 6.4))
    for r, (path, label) in enumerate(cases):
        gray = load_gray(path, max_dim=560)
        for c, (title, fn) in enumerate(columns):
            imshow(axes[r, c], fn(gray))
            if r == 0:
                add_subtitle(axes[r, c], title, weight="bold" if title == "DE-Cividis" else None)
        fig.text(0.085, 0.77 - r * 0.295, label, ha="right", va="center",
                 fontsize=10.5, weight="bold")
    plt.subplots_adjust(left=0.11, right=0.995, top=0.91, bottom=0.04, wspace=0.04, hspace=0.16)
    save_figure("fig5_typical_cases.png")


def fig6_ablation() -> None:
    """Figure 6: explain why each DE-Cividis step is kept."""
    gray = load_gray(DATA / "grayscale" / "xray_synthetic.png", max_dim=460)
    columns = [
        ("原灰度", gray_to_rgb(gray)),
        ("完整 DE-Cividis", de_cividis(gray)),
        ("去掉细节增强", de_cividis_no_de(gray)),
        ("去掉亮度回写", de_cividis_no_lw(gray)),
        ("改用 Jet 色图", de_jet_no_cividis(gray)),
    ]
    rows = [
        ("正常视觉", lambda rgb: rgb),
        ("deutan 模拟", lambda rgb: simulate_cvd(rgb, "deutan")),
    ]

    fig, axes = plt.subplots(len(rows), len(columns), figsize=(11.2, 4.9))
    for r, (row_name, view_fn) in enumerate(rows):
        for c, (title, img) in enumerate(columns):
            imshow(axes[r, c], view_fn(img))
            if r == 0:
                add_subtitle(axes[r, c], title, weight="bold" if c == 1 else None)
        fig.text(0.10, 0.71 - r * 0.41, row_name, ha="right", va="center",
                 fontsize=10.5, weight="bold")
    plt.subplots_adjust(left=0.13, right=0.995, top=0.90, bottom=0.05, wspace=0.04, hspace=0.10)
    save_figure("fig6_ablation.png")


def fig7_kodak_overview() -> None:
    """Figure 7: overview of the 24 supplementary Kodak images."""
    items = [(name, gray) for dataset, category, name, gray in collect_images() if dataset == "external_kodak24"]
    items = sorted(items, key=lambda x: x[0])
    ncols = 8
    nrows = int(np.ceil(len(items) / ncols))
    fig, axes = plt.subplots(nrows, ncols, figsize=(12.4, 4.7))
    axes = np.asarray(axes).reshape(nrows, ncols)
    for ax in axes.flat:
        ax.axis("off")
    for i, (name, gray) in enumerate(items):
        ax = axes.flat[i]
        imshow(ax, de_cividis(gray))
        ax.set_title(name.replace("kodim", "Kodak "), fontsize=7.0, pad=2)
    plt.subplots_adjust(left=0.02, right=0.995, top=0.94, bottom=0.04, wspace=0.03, hspace=0.18)
    save_figure("fig7_kodak_overview.png")



def fig9_metric_bars() -> None:
    """Supplementary metric overview. It is generated for traceability only."""
    summary_path = RESULTS / "final_results_summary.csv"
    if not summary_path.exists():
        print("skip fig9_supp_metric_bars.png: final_results_summary.csv not found")
        return

    grouped: dict[tuple[str, str, str], dict[str, str]] = {}
    with open(summary_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            grouped[(row["dataset"], row["method"], row["cvd_type"])] = row

    methods = [
        ("Jet", "M4_jet"),
        ("Turbo", "M5b_turbo"),
        ("Viridis", "M6_viridis"),
        ("Cividis", "M6b_cividis"),
        ("DE-Cividis", "DE_Cividis"),
    ]
    labels = [m[0] for m in methods]
    avg_grad, de, yssim = [], [], []
    for _, method in methods:
        rows = [grouped[("primary27", method, cvd)] for cvd in ("protan", "deutan")]
        avg_grad.append(np.mean([float(r["avg_gradient_mean"]) for r in rows]))
        de.append(np.mean([float(r["delta_e2000_mean"]) for r in rows]))
        yssim.append(np.mean([float(r["ssim_mean"]) for r in rows]))

    x = np.arange(len(labels))
    fig, axes = plt.subplots(1, 3, figsize=(10.5, 3.1))
    colors = ["#94A3B8", "#F59E0B", "#10B981", "#3B82F6", "#22C55E"]
    specs = [
        (avg_grad, "AvgGrad ↑", (0, max(avg_grad) + 0.08)),
        (de, "ΔE2000 ↓", (10, max(de) + 2.5)),
        (yssim, "Y-SSIM ↑", (0.15, 1.005)),
    ]
    for ax, (vals, title, ylim) in zip(axes, specs):
        ax.bar(x, vals, color=colors, edgecolor="#334155", linewidth=0.8)
        ax.set_title(title, fontsize=10.5, weight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=20, ha="right", fontsize=8.5)
        ax.set_ylim(*ylim)
        ax.grid(axis="y", color="#E5E7EB", linewidth=0.7)
        ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    save_figure("fig9_supp_metric_bars.png")


def main() -> None:
    FIGS.mkdir(exist_ok=True)
    fig1_pipeline()
    fig2_colormap_cvd()
    fig3_parameter_effect()
    fig4_evaluation_overview()
    fig5_typical_cases()
    fig6_ablation()
    fig7_kodak_overview()


if __name__ == "__main__":
    main()
