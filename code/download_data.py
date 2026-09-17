"""Download remaining test images with proper User-Agent."""
import urllib.request
import ssl
import os
import time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# Wikimedia requires identifying User-Agent per their robot policy
WIKI_UA = "ColorResearchBot/1.0 (figure reproduction; research use)"


def dl(path, url, ua=None, sleep_after=2):
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        print(f"SKIP {path} (exists, {os.path.getsize(path)} bytes)")
        return True
    h = {"User-Agent": ua or "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers=h)
            with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
                data = r.read()
                if len(data) > 1000:
                    os.makedirs(os.path.dirname(path), exist_ok=True)
                    open(path, "wb").write(data)
                    print(f"OK   {path}: {len(data)} bytes")
                    time.sleep(sleep_after)
                    return True
                else:
                    print(f"TINY {path}: {len(data)} bytes")
        except Exception as e:
            print(f"  attempt {attempt+1} {path}: {type(e).__name__}: {str(e)[:80]}")
        time.sleep(4)
    return False


# Fix: Lena USC-SIPI — try alt URL
dl("data/usc-sipi/lena.tiff",
   "https://sipi.usc.edu/database/preview/misc/4.2.06.png")
# If that's not Lena, use Wikimedia Lena copy
dl("data/usc-sipi/lena.png",
   "https://upload.wikimedia.org/wikipedia/en/7/7d/Lenna_%28test_image%29.png",
   ua=WIKI_UA)

# Fix Shanghai map tile (different coords)
dl("data/maps/shanghai.png",
   "https://tile.openstreetmap.org/8/215/106.png")
# Backup another colorful map
dl("data/maps/global.png",
   "https://tile.openstreetmap.org/3/4/3.png")

# Ishihara plates with proper UA + delay
ishihara = [
    ("data/ishihara/plate01.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/9/9a/Ishihara_9.svg/512px-Ishihara_9.svg.png"),
    ("data/ishihara/plate02.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/b/b6/Ishihara_11.svg/512px-Ishihara_11.svg.png"),
    ("data/ishihara/plate03.png",
     "https://upload.wikimedia.org/wikipedia/commons/c/c4/Ishihara_23.PNG"),
    ("data/ishihara/plate04.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/0/0b/Ishihara_Plate_3.jpg/512px-Ishihara_Plate_3.jpg"),
    ("data/ishihara/plate05.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e6/Ishihara_8.svg/512px-Ishihara_8.svg.png"),
    ("data/ishihara/plate06.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/2/26/Ishihara_29.svg/512px-Ishihara_29.svg.png"),
]
for path, url in ishihara:
    dl(path, url, ua=WIKI_UA, sleep_after=3)

# Grayscale science visualizations (Wikimedia)
gray = [
    ("data/grayscale/xray_hand.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/d/de/X-ray_of_normal_hand_by_dorsoplantar_projection.jpg/512px-X-ray_of_normal_hand_by_dorsoplantar_projection.jpg"),
    ("data/grayscale/xray_chest.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/9/96/Chest_X-ray_-_pulmonary_oedema.jpg/512px-Chest_X-ray_-_pulmonary_oedema.jpg"),
    ("data/grayscale/sem_pollen.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/5/52/Misc_pollen.jpg/512px-Misc_pollen.jpg"),
    ("data/grayscale/satellite_hurricane.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/Hurricane_Isabel_from_ISS.jpg/512px-Hurricane_Isabel_from_ISS.jpg"),
    ("data/grayscale/weld_radiograph.png",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f3/Radiographs.png/512px-Radiographs.png"),
    ("data/grayscale/mri_brain.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/MRI_brain.jpg/512px-MRI_brain.jpg"),
    ("data/grayscale/moon.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/FullMoon2010.jpg/512px-FullMoon2010.jpg"),
    ("data/grayscale/microscopy.jpg",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/3/35/Cell_culture.jpg/512px-Cell_culture.jpg"),
]
for path, url in gray:
    dl(path, url, ua=WIKI_UA, sleep_after=3)

# Final report
print("\n=== Download summary ===")
for d in ["data/kodak", "data/usc-sipi", "data/ishihara", "data/maps", "data/grayscale"]:
    if os.path.isdir(d):
        files = [f for f in os.listdir(d) if os.path.getsize(os.path.join(d, f)) > 1000]
        print(f"  {d}: {len(files)} files")
