"""
Generate synthetic grayscale scientific test images that mimic the textbook
examples in Ch6.3 (X-ray, weld, satellite, etc.). These are deterministic and
don't depend on flaky external downloads.

Plus convert a few Kodak images to grayscale for additional variety.
"""
import numpy as np
from PIL import Image
import os

os.makedirs("data/grayscale", exist_ok=True)


def save_gray(path, arr):
    """Save uint8 grayscale image."""
    arr = np.clip(arr, 0, 255).astype(np.uint8)
    Image.fromarray(arr, mode="L").save(path)
    print(f"  generated {path}: shape={arr.shape}, range=[{arr.min()},{arr.max()}]")


# 1. Synthetic X-ray of an object (radial gradient + dense regions)
H, W = 384, 384
yy, xx = np.mgrid[0:H, 0:W]
cy, cx = H / 2, W / 2
r = np.sqrt((yy - cy) ** 2 + (xx - cx) ** 2)
xray1 = 220 - 0.5 * r
# Add a "denser" rectangular region (like a bone or implant)
xray1[100:200, 130:250] -= 80
xray1[140:170, 160:220] -= 60  # higher density inner
# Add some texture noise
np.random.seed(42)
xray1 += np.random.normal(0, 5, (H, W))
save_gray("data/grayscale/xray_synthetic.png", xray1)

# 2. Welding seam X-ray (linear high-density band with defects)
weld = np.full((256, 512), 180, dtype=float)
# horizontal weld band
weld[100:160, :] = 80
# Some "defects" (porosity, cracks)
np.random.seed(11)
defect_locs = np.random.randint(0, 480, 12)
for x in defect_locs:
    weld[120 + np.random.randint(-15, 15), x:x + np.random.randint(3, 12)] = 230
weld += np.random.normal(0, 6, weld.shape)
save_gray("data/grayscale/weld_synthetic.png", weld)

# 3. Satellite cloud-like grayscale (smoothed perlin-like noise)
np.random.seed(7)
sat = np.random.normal(128, 40, (256, 256))
# Smooth multiple times
from scipy.ndimage import gaussian_filter
sat = gaussian_filter(sat, sigma=8)
sat = (sat - sat.min()) / (sat.max() - sat.min()) * 255
# Add a swirl pattern (like hurricane)
yy, xx = np.mgrid[0:256, 0:256]
cy, cx = 128, 128
ang = np.arctan2(yy - cy, xx - cx)
rr = np.sqrt((yy - cy) ** 2 + (xx - cx) ** 2)
swirl = np.sin(ang * 3 + rr * 0.05) * 30 * np.exp(-rr / 100)
sat = sat + swirl
save_gray("data/grayscale/satellite_synthetic.png", sat)

# 4. SEM-like fibrous structure
np.random.seed(33)
sem = np.zeros((256, 256), dtype=float)
for _ in range(40):
    x0, y0 = np.random.randint(0, 256, 2)
    angle = np.random.uniform(0, np.pi)
    length = np.random.randint(50, 200)
    intensity = np.random.uniform(120, 200)
    for t in range(length):
        x = int(x0 + t * np.cos(angle))
        y = int(y0 + t * np.sin(angle))
        if 0 <= x < 256 and 0 <= y < 256:
            sem[y, x] = intensity
            for dx in range(-2, 3):
                for dy in range(-2, 3):
                    if 0 <= x + dx < 256 and 0 <= y + dy < 256:
                        sem[y + dy, x + dx] = max(sem[y + dy, x + dx], intensity * 0.6)
sem += 30 + np.random.normal(0, 8, sem.shape)
save_gray("data/grayscale/sem_synthetic.png", sem)

# 5. Rainfall map (smooth gradient with hot/cold regions, like课件 p41)
np.random.seed(101)
rain = np.zeros((200, 320), dtype=float)
# Several "rain centers"
centers = [(50, 80, 200, 40), (140, 200, 180, 60), (100, 280, 150, 35)]  # y, x, max, sigma
for cy_, cx_, mx, sg in centers:
    yy, xx = np.mgrid[0:200, 0:320]
    rain += mx * np.exp(-((yy - cy_) ** 2 + (xx - cx_) ** 2) / (2 * sg ** 2))
rain = np.clip(rain, 0, 255)
save_gray("data/grayscale/rainfall_synthetic.png", rain)

# 6. MRI-like slice (concentric soft regions)
np.random.seed(50)
mri = np.zeros((256, 256), dtype=float)
yy, xx = np.mgrid[0:256, 0:256]
# brain outline
brain_mask = ((yy - 128) ** 2 / (95 ** 2) + (xx - 128) ** 2 / (110 ** 2)) <= 1
mri[brain_mask] = 130
# white/gray matter
inner_mask = ((yy - 128) ** 2 / (70 ** 2) + (xx - 128) ** 2 / (85 ** 2)) <= 1
mri[inner_mask] = 90
# ventricles
vent_mask = ((yy - 128) ** 2 / (15 ** 2) + (xx - 128) ** 2 / (40 ** 2)) <= 1
mri[vent_mask] = 200
# Smooth + noise
mri = gaussian_filter(mri, sigma=2)
mri += np.random.normal(0, 4, mri.shape)
save_gray("data/grayscale/mri_synthetic.png", mri)

# 7. Lena grayscale (convert from existing Lena PNG)
try:
    lena = np.array(Image.open("data/usc-sipi/lena.png").convert("L"))
    save_gray("data/grayscale/lena_gray.png", lena)
except Exception as e:
    print(f"Lena gray skip: {e}")


# === Generate 6 Ishihara-style synthetic plates (PD by construction) ===
os.makedirs("data/ishihara", exist_ok=True)


def make_ishihara(out_path, target_color_rgb, bg_color_rgb, hidden_pattern, seed=42):
    """Generate a pseudo-Ishihara plate with random circles forming a hidden number."""
    np.random.seed(seed)
    size = 384
    img = np.ones((size, size, 3), dtype=np.uint8) * 240  # light bg

    # Hidden pattern is binary mask H,W
    pattern = np.array(hidden_pattern)
    pH, pW = pattern.shape
    mask = np.zeros((size, size), dtype=bool)
    # Place pattern in center
    sx = (size - pW * 24) // 2
    sy = (size - pH * 24) // 2
    for i in range(pH):
        for j in range(pW):
            if pattern[i, j]:
                mask[sy + i * 24:sy + (i + 1) * 24, sx + j * 24:sx + (j + 1) * 24] = True

    # Add random circles
    n_dots = 1500
    for _ in range(n_dots):
        cx = np.random.randint(0, size)
        cy = np.random.randint(0, size)
        r = np.random.randint(4, 11)
        color = target_color_rgb if mask[cy, cx] else bg_color_rgb
        color = np.array(color) + np.random.randint(-20, 20, 3)
        color = np.clip(color, 0, 255).astype(np.uint8)

        # Draw dot
        for dy in range(-r, r + 1):
            for dx in range(-r, r + 1):
                if dx * dx + dy * dy <= r * r:
                    yy, xx = cy + dy, cx + dx
                    if 0 <= yy < size and 0 <= xx < size:
                        img[yy, xx] = color

    Image.fromarray(img).save(out_path)
    print(f"  generated {out_path}")


# 8x5 patterns for digits 1-6
digit_patterns = {
    1: [
        [0, 0, 1, 0, 0],
        [0, 1, 1, 0, 0],
        [0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0],
        [0, 0, 1, 0, 0],
        [0, 1, 1, 1, 0],
    ],
    2: [
        [0, 1, 1, 1, 0],
        [1, 0, 0, 0, 1],
        [0, 0, 0, 1, 0],
        [0, 0, 1, 0, 0],
        [0, 1, 0, 0, 0],
        [1, 0, 0, 0, 0],
        [1, 1, 1, 1, 1],
    ],
    3: [
        [0, 1, 1, 1, 0],
        [1, 0, 0, 0, 1],
        [0, 0, 0, 0, 1],
        [0, 0, 1, 1, 0],
        [0, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [0, 1, 1, 1, 0],
    ],
    5: [
        [1, 1, 1, 1, 1],
        [1, 0, 0, 0, 0],
        [1, 1, 1, 1, 0],
        [0, 0, 0, 0, 1],
        [0, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [0, 1, 1, 1, 0],
    ],
    6: [
        [0, 1, 1, 1, 0],
        [1, 0, 0, 0, 0],
        [1, 1, 1, 1, 0],
        [1, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [0, 1, 1, 1, 0],
    ],
    8: [
        [0, 1, 1, 1, 0],
        [1, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [0, 1, 1, 1, 0],
        [1, 0, 0, 0, 1],
        [1, 0, 0, 0, 1],
        [0, 1, 1, 1, 0],
    ],
}

# Red/green Ishihara-style plates
make_ishihara("data/ishihara/plate01_synthetic.png",
              target_color_rgb=(220, 80, 80),    # red figure
              bg_color_rgb=(120, 180, 100),      # green bg
              hidden_pattern=digit_patterns[1], seed=1)
make_ishihara("data/ishihara/plate02_synthetic.png",
              target_color_rgb=(200, 100, 80),
              bg_color_rgb=(140, 170, 80),
              hidden_pattern=digit_patterns[2], seed=2)
make_ishihara("data/ishihara/plate03_synthetic.png",
              target_color_rgb=(220, 90, 90),
              bg_color_rgb=(130, 175, 90),
              hidden_pattern=digit_patterns[3], seed=3)
make_ishihara("data/ishihara/plate05_synthetic.png",
              target_color_rgb=(210, 100, 110),
              bg_color_rgb=(150, 180, 110),
              hidden_pattern=digit_patterns[5], seed=5)
make_ishihara("data/ishihara/plate06_synthetic.png",
              target_color_rgb=(220, 90, 70),
              bg_color_rgb=(130, 170, 100),
              hidden_pattern=digit_patterns[6], seed=6)
make_ishihara("data/ishihara/plate08_synthetic.png",
              target_color_rgb=(215, 95, 105),
              bg_color_rgb=(135, 175, 95),
              hidden_pattern=digit_patterns[8], seed=8)

# Final report
print("\n=== Synthetic data summary ===")
for d in ["data/kodak", "data/usc-sipi", "data/ishihara", "data/maps", "data/grayscale"]:
    if os.path.isdir(d):
        files = [f for f in os.listdir(d) if os.path.getsize(os.path.join(d, f)) > 1000]
        print(f"  {d}: {len(files)} files")
total = sum(len([f for f in os.listdir(d) if os.path.getsize(os.path.join(d, f)) > 1000])
            for d in ["data/kodak", "data/usc-sipi", "data/ishihara", "data/maps", "data/grayscale"])
print(f"  TOTAL: {total} usable images")
