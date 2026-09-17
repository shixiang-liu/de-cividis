"""
run_final_experiments.py - final experiment table for the paper.

The paper-facing method list is limited to classic pseudocolor methods, public
colormaps, and DE_Cividis as the only proposed method.
"""
import csv
import os
from pathlib import Path

import cv2
import numpy as np
from PIL import Image

from metrics import all_metrics
from pseudocolor_classic import (
    m1_intensity_slicing,
    m2_sin_three_phase,
    m4_jet,
    m5_hot,
    m5b_turbo,
    m6_viridis,
    m6_cividis,
)
from pseudocolor_proposed import de_cividis


ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "results"
CVD_TYPES = ("protan", "deutan")


METHODS = {
    "M1_intensity_slicing": lambda gray, cvd: m1_intensity_slicing(gray),
    "M2_sin_three_phase": lambda gray, cvd: m2_sin_three_phase(gray),
    "M4_jet": lambda gray, cvd: m4_jet(gray),
    "M5_hot": lambda gray, cvd: m5_hot(gray),
    "M5b_turbo": lambda gray, cvd: m5b_turbo(gray),
    "M6_viridis": lambda gray, cvd: m6_viridis(gray),
    "M6b_cividis": lambda gray, cvd: m6_cividis(gray),
    "DE_Cividis": lambda gray, cvd: de_cividis(gray),
}

EXTERNAL_METHODS = set(METHODS)


def _load_image(path, max_side=600):
    arr = np.array(Image.open(path).convert("RGB"))
    h, w = arr.shape[:2]
    if max(h, w) > max_side:
        scale = max_side / max(h, w)
        arr = cv2.resize(
            arr,
            (max(1, int(round(w * scale))), max(1, int(round(h * scale)))),
            interpolation=cv2.INTER_AREA,
        )
    return cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY).astype(np.float64) / 255.0


def collect_images():
    items = []
    for category in ("usc-sipi", "ishihara", "maps", "grayscale"):
        folder = ROOT / "data" / category
        if not folder.exists():
            continue
        for path in sorted(folder.iterdir()):
            if not path.is_file() or path.stat().st_size < 1000:
                continue
            try:
                items.append(("primary27", category, path.stem, _load_image(path)))
            except Exception as exc:
                print(f"skip {path}: {exc}")

    kodak_paths = []
    for folder in (ROOT / "data" / "kodak", ROOT / "data" / "external_kodak_extra"):
        if folder.exists():
            kodak_paths.extend(sorted(folder.glob("*.png")))
    for path in kodak_paths:
        try:
            items.append(("external_kodak24", "kodak-extra", path.stem, _load_image(path)))
        except Exception as exc:
            print(f"skip {path}: {exc}")
    return items


def run():
    RESULTS.mkdir(exist_ok=True)
    items = collect_images()
    print(f"loaded {len(items)} images")
    for ds in sorted({x[0] for x in items}):
        print(f"  {ds}: {sum(1 for x in items if x[0] == ds)}")

    rows = []
    total = sum(
        (len(METHODS) if dataset == "primary27" else len(EXTERNAL_METHODS)) * len(CVD_TYPES)
        for dataset, _, _, _ in items
    )
    done = 0
    cvd_dependent = set()

    for dataset, category, image, gray in items:
        active_methods = METHODS if dataset == "primary27" else {
            key: fn for key, fn in METHODS.items() if key in EXTERNAL_METHODS
        }
        outputs = {}
        for cvd in CVD_TYPES:
            for method, fn in active_methods.items():
                key = (method, cvd) if method in cvd_dependent else (method,)
                if key in outputs:
                    continue
                outputs[key] = fn(gray, cvd)

        for cvd in CVD_TYPES:
            for method in active_methods:
                key = (method, cvd) if method in cvd_dependent else (method,)
                out = outputs[key]
                metrics = all_metrics(
                    out,
                    gray_ref=gray,
                    cvd_type=cvd,
                )
                rows.append({
                    "dataset": dataset,
                    "category": category,
                    "image": image,
                    "method": method,
                    "cvd_type": cvd,
                    **{k: f"{v:.5f}" for k, v in metrics.items()},
                })
                done += 1
                if done % 100 == 0 or done == total:
                    print(f"  progress {done}/{total}", flush=True)

    out_csv = RESULTS / "final_results.csv"
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    groups = {}
    for row in rows:
        key = (row["dataset"], row["method"], row["cvd_type"])
        groups.setdefault(key, []).append(row)

    out_summary = RESULTS / "final_results_summary.csv"
    fields = [
        "dataset", "method", "cvd_type", "n",
        "delta_e2000_mean", "delta_e2000_std",
        "ssim_mean", "ssim_std",
        "ssim_rgb_mean", "ssim_rgb_std",
        "avg_gradient_mean", "avg_gradient_std",
        "entropy_mean", "entropy_std",
    ]
    with open(out_summary, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for (dataset, method, cvd), rs in sorted(groups.items()):
            vals = {}
            for metric in ("delta_e2000", "ssim", "ssim_rgb", "avg_gradient", "entropy"):
                arr = np.array([float(r[metric]) for r in rs], dtype=np.float64)
                vals[f"{metric}_mean"] = f"{arr.mean():.5f}"
                vals[f"{metric}_std"] = f"{arr.std():.5f}"
            writer.writerow({
                "dataset": dataset,
                "method": method,
                "cvd_type": cvd,
                "n": len(rs),
                **vals,
            })

    print(f"saved {out_csv}")
    print(f"saved {out_summary}")


if __name__ == "__main__":
    run()
