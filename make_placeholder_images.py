"""
각 슬라이드 용도에 맞는 분위기 placeholder 이미지 생성 (PIL 사용)
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
import random, math

Path("images").mkdir(exist_ok=True)

def gradient_image(w, h, c1, c2, angle=0):
    img = Image.new("RGB", (w, h))
    px = img.load()
    rad = math.radians(angle)
    cos_a, sin_a = math.cos(rad), math.sin(rad)
    for y in range(h):
        for x in range(w):
            t = (x * cos_a + y * sin_a) / (w * cos_a + h * sin_a + 1e-9)
            t = max(0, min(1, t))
            r = int(c1[0] + (c2[0] - c1[0]) * t)
            g = int(c1[1] + (c2[1] - c1[1]) * t)
            b = int(c1[2] + (c2[2] - c1[2]) * t)
            px[x, y] = (r, g, b)
    return img

def add_noise(img, amt=12):
    px = img.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            n = random.randint(-amt, amt)
            r, g, b = px[x, y]
            px[x, y] = (max(0,min(255,r+n)), max(0,min(255,g+n)), max(0,min(255,b+n)))
    return img

def add_charcoal_strokes(img, color=(30,30,30), count=8):
    draw = ImageDraw.Draw(img, "RGBA")
    w, h = img.size
    for _ in range(count):
        x0 = random.randint(0, w)
        y0 = random.randint(0, h)
        x1 = x0 + random.randint(-w//3, w//3)
        y1 = y0 + random.randint(-h//3, h//3)
        thick = random.randint(20, 80)
        alpha = random.randint(60, 140)
        draw.line([(x0,y0),(x1,y1)], fill=color+(alpha,), width=thick)
    return img

configs = {
    # img_1: 표지 오른쪽 - 야외 조각, 어두운 하늘/산
    "img_1": dict(c1=(0x1A,0x1A,0x28), c2=(0x08,0x08,0x10), angle=160, strokes=True, stroke_color=(15,15,20), blur=2),
    # img_2: 숯의 언어 배경 - 어두운 숯 질감
    "img_2": dict(c1=(0x0A,0x0A,0x0A), c2=(0x1E,0x1E,0x1E), angle=90, strokes=True, stroke_color=(5,5,5), blur=1),
    # img_3: 엔딩 소형 이미지 - 갤러리 내부
    "img_3": dict(c1=(0x22,0x22,0x22), c2=(0x10,0x10,0x10), angle=45, strokes=False, blur=0),
    # img_4: 기다림 배경 - 영상 설치, 어둡고 신비
    "img_4": dict(c1=(0x08,0x08,0x10), c2=(0x18,0x14,0x20), angle=270, strokes=True, stroke_color=(20,18,25), blur=3),
    # img_5: 전시 공간 배경 - 붓질 연작, 밝은 콘크리트
    "img_5": dict(c1=(0x2A,0x2A,0x2A), c2=(0x14,0x14,0x14), angle=30, strokes=True, stroke_color=(50,50,50), blur=1),
    # img_6: 엔딩 메인 이미지 - 조각 대형
    "img_6": dict(c1=(0x10,0x10,0x10), c2=(0x28,0x25,0x1A), angle=120, strokes=True, stroke_color=(20,18,10), blur=2),
    # img_7: 작가 소개 오른쪽 - 작가 포트레이트 느낌
    "img_7": dict(c1=(0x18,0x14,0x10), c2=(0x08,0x08,0x08), angle=200, strokes=False, blur=1),
}

W, H = 1200, 900
for key, cfg in configs.items():
    img = gradient_image(W, H, cfg["c1"], cfg["c2"], cfg.get("angle",0))
    if cfg.get("strokes"):
        img = add_charcoal_strokes(img, color=cfg.get("stroke_color",(20,20,20)), count=10)
    img = add_noise(img, 8)
    if cfg.get("blur",0):
        img = img.filter(ImageFilter.GaussianBlur(cfg["blur"]))
    out = Path("images") / f"{key}.jpg"
    img.save(out, "JPEG", quality=88)
    print(f"  생성: {out}  ({out.stat().st_size:,}B)")

print("완료!")
