"""
IMAGE RESIZER BY SIAR DIGITAL — Advanced Batch Engine v4.0
Created by Asif Nawaz | Siar Digital 2026
"""

import os

import sys

import gc

import threading

import queue

import time

import logging

import traceback

import csv

import io

from concurrent.futures import ThreadPoolExecutor, as_completed

from datetime import datetime

from pathlib import Path

# ── Core UI ──────────────────────────────────────────────────────────────────

try:

    import customtkinter as ctk

except ImportError:

    print("ERROR: customtkinter not found.\nInstall: pip install customtkinter")

    sys.exit(1)

try:

    from PIL import Image, ImageTk, ImageOps, ImageFilter, ImageDraw, ImageEnhance

except ImportError:

    print("ERROR: Pillow not found.\nInstall: pip install Pillow")

    sys.exit(1)

try:

    from tkinter import filedialog, messagebox

    import tkinter as tk

except ImportError:

    print("ERROR: tkinter not found.")

    sys.exit(1)

# ── Optional Power Libraries ─────────────────────────────────────────────────

try:

    import cv2

    import numpy as np

    HAS_OPENCV = True

except ImportError:

    HAS_OPENCV = False

    cv2 = None

    np = None

try:

    from ultralytics import YOLO

    HAS_YOLO = True

    _yolo_model = None

except ImportError:

    HAS_YOLO = False

    YOLO = None

    _yolo_model = None

try:

    from rembg import remove as rembg_remove

    HAS_REMBG = True

except ImportError:

    HAS_REMBG = False

    rembg_remove = None

try:

    import mediapipe as mp

    HAS_MEDIAPIPE = True

except ImportError:

    HAS_MEDIAPIPE = False

    mp = None

try:

    import openpyxl

    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    HAS_OPENPYXL = True

except ImportError:

    HAS_OPENPYXL = False

    openpyxl = None

try:

    import psutil

    HAS_PSUTIL = True

except ImportError:

    HAS_PSUTIL = False

# ── Logging ───────────────────────────────────────────────────────────────────

log_dir = Path.home() / "SiarDigital_ImageResizer_Logs"

log_dir.mkdir(exist_ok=True)

log_file = log_dir / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(

    level=logging.INFO,

    format="%(asctime)s [%(levelname)s] %(message)s",

    handlers=[

        logging.FileHandler(log_file, encoding="utf-8"),

        logging.StreamHandler(sys.stdout),

    ],

)

log = logging.getLogger("SiarDigital_Resizer")

# ── Design Tokens ─────────────────────────────────────────────────────────────

C = {

    "bg":       "#060910",

    "surface":  "#0D1321",

    "surface2": "#131D2E",

    "surface3": "#1A2540",

    "border":   "#1E2D45",

    "border2":  "#2D4060",

    "accent":   "#00F5A0",   # neon green

    "accent2":  "#00D4FF",   # cyan

    "accent3":  "#FF4D6D",   # red

    "accent4":  "#FFD166",   # amber

    "accent5":  "#B388FF",   # violet

    "accent6":  "#FF9F43",   # orange

    "accent7":  "#06D6A0",   # teal

    "text":     "#E8F0FF",

    "text2":    "#7A95C0",

    "text3":    "#3A4F70",

    "success":  "#00F5A0",

    "error":    "#FF4D6D",

    "warning":  "#FFD166",

    "gold":     "#FFD700",

}

FONTS = {

    "hero":       ("Consolas", 20, "bold"),

    "title":      ("Consolas", 14, "bold"),

    "subtitle":   ("Consolas", 9),

    "label":      ("Consolas", 10, "bold"),

    "body":       ("Consolas", 11),

    "small":      ("Consolas", 9),

    "tiny":       ("Consolas", 8),

    "big_btn":    ("Consolas", 13, "bold"),

    "counter":    ("Consolas", 24, "bold"),

    "credit":     ("Consolas", 13, "bold"),

    "credit_sm":  ("Consolas", 9),

    "badge":      ("Consolas", 8, "bold"),

}

SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.webp', '.tiff', '.tif',

                  '.bmp', '.gif', '.ico', '.ppm')

QUALITY_LIMIT_KB = 1536   # 1.5 MB quality guard

# ── Presets ───────────────────────────────────────────────────────────────────

PRESETS = {

    "Instagram Post (1:1)":    (1080, 1080),

    "Instagram Story (9:16)":  (1080, 1920),

    "Instagram Portrait (4:5)":(1080, 1350),

    "Facebook Cover":          (820,  312),

    "Twitter Header":          (1500, 500),

    "YouTube Thumbnail (16:9)":(1280, 720),

    "LinkedIn Banner":         (1584, 396),

    "WhatsApp DP":             (800,  800),

    "TikTok Video Cover":      (1080, 1920),

    "Amazon Product":          (2000, 2000),

    "4K Wallpaper":            (3840, 2160),

    "Full HD Wallpaper":       (1920, 1080),

    "A4 Print (300 DPI)":      (2480, 3508),

    "Custom":                  None,

}

# ── Category Modes ────────────────────────────────────────────────────────────

CATEGORIES = {

    "General":   {"icon": "⚙️",  "desc": "Smart multi-purpose detection"},

    "Clothing":  {"icon": "👗",  "desc": "YOLOv8 focuses on human/model posture"},

    "Jewelry":   {"icon": "💍",  "desc": "High-detail extraction for small objects"},

    "Furniture": {"icon": "🪑",  "desc": "Boundary detection for large items"},

    "Portrait":  {"icon": "🧑",  "desc": "Face-priority headroom rules apply"},

    "Product":   {"icon": "📦",  "desc": "AI centers product, fills background"},

}

# ── Placement Modes ──────────────────────────────────────────────────────────

MODE_TIPS = {

    "Smart Crop (AI)":         "AI detects subject, crops to center it perfectly.",

    "Mirror BG (Smart Fill)":  "Fills gaps with blurred mirror of image. No black bars!",

    "AI Background Extend":    "Reconstructs/extends background to fill canvas.",

    "Fill & Crop (Center)":    "Center-crop. No empty space.",

    "Letterbox (Dark BG)":     "Black bars preserve full image.",

    "Letterbox (White BG)":    "White bars for e-commerce.",

    "Stretch to Fit":          "Distorts to exact size.",

}

msg_queue = queue.Queue()

# ─────────────────────────────────────────────────────────────────────────────

#  AI Detection Helpers  —  Human-Eye Algorithm

#  Strategy: detect face at calibrated resolution → GrabCut body extent →

#  compute head_top / feet_bottom / body sides → equal-padding crop box

# ─────────────────────────────────────────────────────────────────────────────

# Fixed detection canvas: tall enough for full-body fashion shots

_DETECT_W = 800

_DETECT_H = 1200

def _get_yolo():

    global _yolo_model

    if HAS_YOLO and _yolo_model is None:

        try:

            _yolo_model = YOLO("yolov8n.pt")

            log.info("YOLOv8 model loaded.")

        except Exception as e:

            log.warning(f"YOLOv8 load failed: {e}")

    return _yolo_model

def _detect_face_cv(gray_img, sw, sh):

    """

    Robust face detection calibrated for high-res fashion/portrait photography.

    Strategy:

    1. Try scale_factor=1.1 first (tight, reduces false positives).

    2. Fall back to scale_factor=1.05 if 1.1 finds nothing.

    3. Score each candidate:

         pos_score: prefer faces at 20-70% of image height (avoid wall noise)

         size_score: prefer larger faces

    4. Profile cascade tried only if frontal finds nothing.

    Returns (fx, fy, fw, fh) or None.

    """

    cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"

    profile_path = cv2.data.haarcascades + "haarcascade_profileface.xml"

    all_candidates = []   # (fx, fy, fw, fh, score)

    for cpath in [cascade_path, profile_path]:

        if not os.path.exists(cpath):

            continue

        fc = cv2.CascadeClassifier(cpath)

        # strict scale first (fewer false positives), then looser as fallback

        for scale_f in [1.1, 1.05]:

            for min_nb in [3, 4, 5]:

                try:

                    faces = fc.detectMultiScale(

                        gray_img, scale_f, min_nb, minSize=(12, 12))

                except Exception:

                    continue

                for fx, fy, fw, fh in faces:

                    if fh < sh * 0.02:           # skip tiny false positives

                        continue

                    if fy < sh * 0.03:           # skip top-edge wall noise

                        continue

                    y_center_frac = (fy + fh / 2) / sh

                    # Gaussian-like weight centred at 40-65% of canvas height

                    pos_score  = max(0.0, 1.0 - abs(y_center_frac - 0.525) * 2.8)

                    size_score = min(1.0, (fh / sh) * 10)

                    score = pos_score * size_score

                    all_candidates.append((fx, fy, fw, fh, score))

        if all_candidates:

            break   # frontal found something; skip profile

    if not all_candidates:

        return None

    best = max(all_candidates, key=lambda c: c[4])

    return (best[0], best[1], best[2], best[3])

def _detect_person_bbox_cv(img_pil, category):

    """

    Full human-eye pipeline:

    1. Resize to calibrated canvas (_DETECT_W x _DETECT_H)

    2. Face detection (multi-cascade, multi-param sweep)

    3. GrabCut seeded from face → body extent

    4. Scale bbox back to original pixel coords

    Returns dict or None.

    """

    iw, ih = img_pil.size

    # ── Resize to detection canvas ────────────────────────────────────────────

    small = img_pil.resize((_DETECT_W, _DETECT_H), Image.Resampling.LANCZOS)

    cv_img = cv2.cvtColor(np.array(small), cv2.COLOR_RGB2BGR)

    gray   = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)

    sh, sw = _DETECT_H, _DETECT_W

    # ── Scale factors ─────────────────────────────────────────────────────────

    sx = iw / sw   # multiply detection px → original px

    sy = ih / sh

    # ── Step 1: Face detection ────────────────────────────────────────────────

    face = _detect_face_cv(gray, sw, sh)

    if face is None:

        log.info("No face detected — falling back to saliency center.")

        return None

    fx, fy, fw, fh = face

    log.info(f"Face detected @ ({fx},{fy}) {fw}x{fh} in {sw}x{sh} canvas "

             f"[top={fy/sh:.1%}  left={fx/sw:.1%}]")

    # ── Step 2: GrabCut seeded from face region ───────────────────────────────

    margin   = int(fw * 1.0)

    rect_x   = max(0, fx - margin)

    rect_y   = max(0, fy - int(fh * 0.4))          # include hair

    rect_w   = min(sw - rect_x, fw + 2 * margin)

    rect_h   = min(sh - rect_y, sh - rect_y)        # extend to bottom

    body_left_s   = rect_x

    body_right_s  = rect_x + rect_w

    body_bottom_s = sh   # default = full image height

    try:

        mask = np.zeros(cv_img.shape[:2], np.uint8)

        bgd  = np.zeros((1, 65), np.float64)

        fgd  = np.zeros((1, 65), np.float64)

        cv2.grabCut(cv_img, mask,

                    (rect_x, rect_y, rect_w, rect_h),

                    bgd, fgd, 5, cv2.GC_INIT_WITH_RECT)

        fg     = np.where((mask == 2) | (mask == 0), 0, 1).astype("uint8")

        coords = np.column_stack(np.where(fg > 0))

        if len(coords) > 200:

            y_min, x_min = coords.min(axis=0)

            y_max, x_max = coords.max(axis=0)

            body_left_s   = x_min

            body_right_s  = x_max

            body_bottom_s = y_max

            log.info(f"GrabCut body: x={x_min}-{x_max}  y={y_min}-{y_max}")

    except Exception as e:

        log.warning(f"GrabCut failed: {e}")

    # ── Step 3: Head top (face top minus hair allowance) ──────────────────────

    head_top_s = max(0, fy - int(fh * 0.65))   # 65% of face-height for hair

    # ── Step 4: Scale everything back to original image coords ───────────────

    def u(v, factor): return int(v * factor)

    return {

        "head_top":    u(head_top_s,    sy),

        "feet_bottom": u(body_bottom_s, sy),

        "body_left":   u(body_left_s,   sx),

        "body_right":  u(body_right_s,  sx),

        "face_cx":     u(fx + fw // 2,  sx),

        "face_cy":     u(fy + fh // 2,  sy),

        "face_w":      u(fw, sx),

        "face_h":      u(fh, sy),

        "img_w":       iw,

        "img_h":       ih,

    }

def _detect_person_bbox_yolo(img_pil, category):

    """YOLOv8 person detection — returns same dict schema or None."""

    model = _get_yolo()

    if model is None or not HAS_OPENCV:

        return None

    try:

        iw, ih = img_pil.size

        # Run YOLO at reduced size for speed

        run_img = img_pil.copy()

        run_img.thumbnail((1280, 1280), Image.Resampling.LANCZOS)

        rw, rh = run_img.size

        sx, sy = iw / rw, ih / rh

        cv_img  = cv2.cvtColor(np.array(run_img), cv2.COLOR_RGB2BGR)

        results = model(cv_img, verbose=False)

        cat_classes = {

            "Clothing":  [0],

            "Portrait":  [0],

            "Jewelry":   list(range(80)),

            "Furniture": [56, 57, 58, 59, 60, 61, 62],

            "Product":   list(range(80)),

            "General":   list(range(80)),

        }

        allowed = cat_classes.get(category, list(range(80)))

        boxes = []

        for r in results:

            for box in r.boxes:

                cls  = int(box.cls[0])

                conf = float(box.conf[0])

                if cls in allowed and conf > 0.25:

                    x1, y1, x2, y2 = box.xyxy[0].tolist()

                    boxes.append((x1, y1, x2, y2, conf))

        if not boxes:

            return None

        # Largest confident box

        boxes.sort(key=lambda b: (b[2]-b[0]) * (b[3]-b[1]) * b[4], reverse=True)

        bx1, by1, bx2, by2, _ = boxes[0]

        # Scale to original

        orig_x1 = int(bx1 * sx);  orig_y1 = int(by1 * sy)

        orig_x2 = int(bx2 * sx);  orig_y2 = int(by2 * sy)

        face_h_est = int((orig_y2 - orig_y1) * 0.12)

        return {

            "head_top":    max(0, orig_y1),

            "feet_bottom": min(ih, orig_y2),

            "body_left":   max(0, orig_x1),

            "body_right":  min(iw, orig_x2),

            "face_cx":     (orig_x1 + orig_x2) // 2,

            "face_cy":     orig_y1 + face_h_est,

            "face_w":      orig_x2 - orig_x1,

            "face_h":      face_h_est,

            "img_w":       iw,

            "img_h":       ih,

        }

    except Exception as e:

        log.warning(f"YOLO detection error: {e}")

        return None

def _detect_person_bbox_saliency(img_pil):

    """

    Pure OpenCV saliency fallback (no face cascade).

    Uses edge density + color saturation to find subject cluster.

    """

    iw, ih = img_pil.size

    if not HAS_OPENCV:

        return None

    try:

        small = img_pil.resize((_DETECT_W, _DETECT_H), Image.Resampling.LANCZOS)

        cv_img = cv2.cvtColor(np.array(small), cv2.COLOR_RGB2BGR)

        gray   = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)

        sh, sw = _DETECT_H, _DETECT_W

        sx, sy = iw / sw, ih / sh

        # Edge map

        blur   = cv2.GaussianBlur(gray, (7, 7), 0)

        edges  = cv2.Canny(blur, 30, 120)

        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15))

        dilated = cv2.dilate(edges, kernel, iterations=3)

        # Saturation map

        hsv = cv2.cvtColor(cv_img, cv2.COLOR_BGR2HSV)

        sat = hsv[:, :, 1]

        combined = cv2.bitwise_or(dilated, (sat > 50).astype(np.uint8) * 255)

        contours, _ = cv2.findContours(combined, cv2.RETR_EXTERNAL,

                                        cv2.CHAIN_APPROX_SIMPLE)

        min_area = sh * sw * 0.005

        boxes = []

        for c in contours:

            area = cv2.contourArea(c)

            if area > min_area:

                x, y, bw, bh = cv2.boundingRect(c)

                boxes.append((x, y, bw, bh, area))

        if not boxes:

            return None

        total_w = sum(b[4] for b in boxes)

        if total_w == 0:

            return None

        # Weighted center

        cx_s = sum((b[0] + b[2]//2) * b[4] for b in boxes) / total_w

        cy_s = sum((b[1] + b[3]//2) * b[4] for b in boxes) / total_w

        # Weighted bbox

        x1_s = min(b[0] for b in boxes)

        y1_s = min(b[1] for b in boxes)

        x2_s = max(b[0] + b[2] for b in boxes)

        y2_s = max(b[1] + b[3] for b in boxes)

        return {

            "head_top":    int(y1_s * sy),

            "feet_bottom": int(y2_s * sy),

            "body_left":   int(x1_s * sx),

            "body_right":  int(x2_s * sx),

            "face_cx":     int(cx_s * sx),

            "face_cy":     int(cy_s * sy),

            "face_w":      0,

            "face_h":      0,

            "img_w":       iw,

            "img_h":       ih,

        }

    except Exception as e:

        log.warning(f"Saliency detection failed: {e}")

        return None

def detect_person_full(img_pil, category="General"):

    """

    Master detector: YOLO → OpenCV face+GrabCut → saliency → None.

    Returns bbox dict or None.

    """

    # 1. YOLO (most accurate)

    bbox = _detect_person_bbox_yolo(img_pil, category)

    if bbox:

        log.info("Using YOLO bbox.")

        return bbox

    # 2. OpenCV face + GrabCut body (calibrated for fashion/portrait)

    if HAS_OPENCV:

        bbox = _detect_person_bbox_cv(img_pil, category)

        if bbox:

            log.info("Using CV face+GrabCut bbox.")

            return bbox

        # 3. Saliency-only fallback

        bbox = _detect_person_bbox_saliency(img_pil)

        if bbox:

            log.info("Using saliency bbox.")

            return bbox

    return None

# ─────────────────────────────────────────────────────────────────────────────

#  Placement Engines

# ─────────────────────────────────────────────────────────────────────────────

def engine_smart_crop(img, tw, th, category="General", **_):

    """

    Human-eye crop:

    • Detect person head_top → feet_bottom → body sides

    • Add equal 10% headroom / 5% footroom / 5% side padding

    • Enforce target aspect ratio by expanding the narrower axis

    • Clamp to image bounds, then scale to exact target size

    Falls back to center-crop if detection fails.

    """

    iw, ih = img.size

    bbox = detect_person_full(img, category)

    if bbox is None:

        # Graceful fallback: center crop

        log.info("Smart crop fallback: center crop.")

        return ImageOps.fit(img, (tw, th), Image.Resampling.LANCZOS,

                            centering=(0.5, 0.5))

    head_top    = bbox["head_top"]

    feet_bottom = bbox["feet_bottom"]

    body_left   = bbox["body_left"]

    body_right  = bbox["body_right"]

    # Sanity clamp

    head_top    = max(0, min(head_top,    ih - 1))

    feet_bottom = max(head_top + 1, min(feet_bottom, ih))

    body_left   = max(0, min(body_left,   iw - 1))

    body_right  = max(body_left + 1, min(body_right, iw))

    person_h = feet_bottom - head_top

    person_w = body_right  - body_left

    if person_h < 10 or person_w < 10:

        log.warning("Detected bbox too small — center crop fallback.")

        return ImageOps.fit(img, (tw, th), Image.Resampling.LANCZOS,

                            centering=(0.5, 0.5))

    # ── Equal padding ─────────────────────────────────────────────────────────

    headroom = int(person_h * 0.10)   # 10% above head

    footroom = int(person_h * 0.05)   # 5%  below feet

    side_pad = int(person_w * 0.05)   # 5%  each side

    crop_top    = head_top    - headroom

    crop_bottom = feet_bottom + footroom

    crop_left   = body_left   - side_pad

    crop_right  = body_right  + side_pad

    crop_h = crop_bottom - crop_top

    crop_w = crop_right  - crop_left

    # ── Enforce target aspect ratio (expand the shorter axis) ────────────────

    target_ratio   = tw / th

    current_ratio  = crop_w / max(crop_h, 1)

    if current_ratio > target_ratio:

        # Too wide → add height equally top & bottom

        new_h  = int(crop_w / target_ratio)

        extra  = new_h - crop_h

        crop_top    -= extra // 2

        crop_bottom += extra - extra // 2

    else:

        # Too tall → add width equally left & right

        new_w  = int(crop_h * target_ratio)

        extra  = new_w - crop_w

        crop_left  -= extra // 2

        crop_right += extra - extra // 2

    # ── Clamp to image bounds ─────────────────────────────────────────────────

    crop_top    = max(0, crop_top)

    crop_bottom = min(ih, crop_bottom)

    crop_left   = max(0, crop_left)

    crop_right  = min(iw, crop_right)

    if crop_right <= crop_left or crop_bottom <= crop_top:

        log.warning("Crop box collapsed — center crop fallback.")

        return ImageOps.fit(img, (tw, th), Image.Resampling.LANCZOS,

                            centering=(0.5, 0.5))

    log.info(f"Smart crop box: ({crop_left},{crop_top})->"

             f"({crop_right},{crop_bottom})  "

             f"[head@{head_top/ih:.1%}  feet@{feet_bottom/ih:.1%}  "

             f"left@{body_left/iw:.1%}  right@{body_right/iw:.1%}]")

    cropped = img.crop((crop_left, crop_top, crop_right, crop_bottom))

    return cropped.resize((tw, th), Image.Resampling.LANCZOS)

def engine_mirror_bg(img, tw, th, **_):

    bg = ImageOps.fit(img.copy(), (tw, th), Image.Resampling.LANCZOS)

    bg = bg.filter(ImageFilter.GaussianBlur(radius=28))

    # Darken blend

    dark = Image.new("RGB", (tw, th), (0, 0, 0))

    bg = Image.blend(bg, dark, alpha=0.35)

    thumb = img.copy()

    thumb.thumbnail((tw, th), Image.Resampling.LANCZOS)

    ox = (tw - thumb.width) // 2

    oy = (th - thumb.height) // 2

    bg.paste(thumb, (ox, oy))

    return bg

def engine_ai_bg_extend(img, tw, th, **_):

    """Extend background using tiled+blurred mirror fill."""

    orig_w, orig_h = img.size

    # Scale original to fit within target keeping aspect

    scale = min(tw / orig_w, th / orig_h)

    nw, nh = int(orig_w * scale), int(orig_h * scale)

    scaled_orig = img.resize((nw, nh), Image.Resampling.LANCZOS)

    # Create extended canvas using mirror tiling

    canvas = Image.new("RGB", (tw, th), (20, 20, 30))

    # Fill with blurred/flipped tiles of the original

    flipped_h = ImageOps.mirror(img)

    flipped_v = ImageOps.flip(img)

    flipped_hv = ImageOps.mirror(flipped_v)

    bg_large = Image.new("RGB", (tw * 2, th * 2))

    for dx, src in [(0, img), (tw, flipped_h)]:

        for dy, s in [(0, src), (th, ImageOps.flip(src))]:

            tile = ImageOps.fit(s, (tw, th), Image.Resampling.LANCZOS)

            bg_large.paste(tile, (dx, dy))

    bg = bg_large.crop((tw//4, th//4, tw//4+tw, th//4+th))

    bg = bg.filter(ImageFilter.GaussianBlur(radius=22))

    dark = Image.new("RGB", (tw, th), (10, 12, 20))

    bg = Image.blend(bg, dark, alpha=0.4)

    # Paste scaled original centered

    ox = (tw - nw) // 2

    oy = (th - nh) // 2

    bg.paste(scaled_orig, (ox, oy))

    return bg

def engine_fill_crop(img, tw, th, **_):

    return ImageOps.fit(img, (tw, th), Image.Resampling.LANCZOS,

                        centering=(0.5, 0.5))

def engine_letterbox_dark(img, tw, th, **_):

    canvas = Image.new("RGB", (tw, th), (12, 14, 20))

    img.thumbnail((tw, th), Image.Resampling.LANCZOS)

    canvas.paste(img, ((tw - img.width) // 2, (th - img.height) // 2))

    return canvas

def engine_letterbox_white(img, tw, th, **_):

    canvas = Image.new("RGB", (tw, th), (255, 255, 255))

    img.thumbnail((tw, th), Image.Resampling.LANCZOS)

    canvas.paste(img, ((tw - img.width) // 2, (th - img.height) // 2))

    return canvas

def engine_stretch(img, tw, th, **_):

    return img.resize((tw, th), Image.Resampling.LANCZOS)

MODE_MAP = {

    "Smart Crop (AI)":         engine_smart_crop,

    "Mirror BG (Smart Fill)":  engine_mirror_bg,

    "AI Background Extend":    engine_ai_bg_extend,

    "Fill & Crop (Center)":    engine_fill_crop,

    "Letterbox (Dark BG)":     engine_letterbox_dark,

    "Letterbox (White BG)":    engine_letterbox_white,

    "Stretch to Fit":          engine_stretch,

}

# ─────────────────────────────────────────────────────────────────────────────

#  Quality Guard (≤ 1.5 MB loop)

# ─────────────────────────────────────────────────────────────────────────────

def save_with_quality_guard(img, out_path, fmt, initial_quality, lossless):

    """Save image, reducing quality until file ≤ 1.5 MB."""

    ext = fmt.lower()

    if lossless:

        img.save(out_path, format="PNG", compress_level=0, optimize=True)

        return

    quality = initial_quality

    while quality >= 30:

        buf = io.BytesIO()

        if ext == "webp":

            img.save(buf, format="WEBP", quality=quality, method=6,

                     lossless=(quality >= 95))

        elif ext in ("jpg", "jpeg"):

            img.save(buf, format="JPEG", quality=quality, optimize=True,

                     progressive=True, subsampling="4:4:4")

        elif ext == "png":

            compress = max(0, min(9, (100 - quality) // 11))

            img.save(buf, format="PNG", compress_level=compress, optimize=True)

        elif ext == "tiff":

            img.save(buf, format="TIFF",

                     compression="lzw" if quality < 90 else "none")

        else:

            img.save(buf, format=ext.upper())

        size_kb = buf.tell() / 1024

        if size_kb <= QUALITY_LIMIT_KB or ext in ("png", "bmp", "tiff"):

            buf.seek(0)

            with open(out_path, "wb") as f:

                f.write(buf.read())

            return

        quality -= 5

    # last resort

    buf.seek(0)

    with open(out_path, "wb") as f:

        f.write(buf.read())

# ─────────────────────────────────────────────────────────────────────────────

#  File Scanner

# ─────────────────────────────────────────────────────────────────────────────

def scan_image_files(input_path):

    input_path = os.path.normpath(os.path.abspath(input_path))

    results = []

    for root, dirs, files in os.walk(input_path):

        dirs[:] = [d for d in dirs if not d.startswith(".")]

        for fname in files:

            if fname.lower().endswith(SUPPORTED_EXTS):

                abs_path = os.path.normpath(os.path.join(root, fname))

                rel_path = os.path.relpath(abs_path, input_path)

                results.append((abs_path, rel_path))

    return results

# ─────────────────────────────────────────────────────────────────────────────

#  Process Single Image

# ─────────────────────────────────────────────────────────────────────────────

def process_one(abs_path, rel_path, out_dir, tw, th, mode, fmt,

                quality, lossless, send_preview, category,

                remove_watermark):

    original_basename = os.path.basename(rel_path)

    if not os.path.isfile(abs_path):

        return (rel_path, "", "", "0.00", "Failed",

                f"Source not found: {abs_path}")

    rel_dir = os.path.dirname(rel_path)

    out_subdir = os.path.join(out_dir, rel_dir) if rel_dir else out_dir

    os.makedirs(out_subdir, exist_ok=True)

    out_path = os.path.join(out_subdir, original_basename)

    try:

        with Image.open(abs_path) as raw:

            try:

                raw = ImageOps.exif_transpose(raw)

            except Exception:

                pass

            img = raw.copy()

        orig_w, orig_h = img.size

        orig_size_kb = os.path.getsize(abs_path) / 1024

        # Normalise mode

        if img.mode in ("P", "PA"):

            img = img.convert("RGBA")

        if img.mode in ("RGBA", "LA"):

            bg = Image.new("RGB", img.size, (255, 255, 255))

            try:

                bg.paste(img, mask=img.split()[-1])

            except Exception:

                bg.paste(img)

            img = bg

        elif img.mode != "RGB":

            img = img.convert("RGB")

        # Optional AI watermark removal via rembg (background-as-watermark)

        if remove_watermark and HAS_REMBG:

            try:

                img_bytes = io.BytesIO()

                img.save(img_bytes, format="PNG")

                img_bytes.seek(0)

                cleaned = rembg_remove(img_bytes.read())

                cleaned_img = Image.open(io.BytesIO(cleaned)).convert("RGBA")

                # Composite on white

                bg2 = Image.new("RGB", cleaned_img.size, (255, 255, 255))

                bg2.paste(cleaned_img, mask=cleaned_img.split()[-1])

                img = bg2

            except Exception as e:

                log.warning(f"Watermark removal failed for {rel_path}: {e}")

        # Preview originals

        p_orig = None

        if send_preview:

            p_orig = img.copy()

            p_orig.thumbnail((240, 180), Image.Resampling.LANCZOS)

        # Apply placement engine

        # Smart Crop passes category for human-eye detection + padding rules.

        engine_fn = MODE_MAP.get(mode, engine_mirror_bg)

        if engine_fn == engine_smart_crop:

            result = engine_fn(img, tw, th, category=category)

        else:

            result = engine_fn(img, tw, th)

        del img

        gc.collect()

        # Preview result

        p_res = None

        if send_preview:

            p_res = result.copy()

            p_res.thumbnail((240, 180), Image.Resampling.LANCZOS)

        # Save with quality guard

        save_with_quality_guard(result, out_path, fmt, quality, lossless)

        resized_size_kb = os.path.getsize(out_path) / 1024

        if send_preview and p_orig and p_res:

            msg_queue.put(("preview", p_orig, p_res,

                           original_basename, resized_size_kb))

        log.info(f"  OK {rel_path}: {orig_w}x{orig_h} -> {tw}x{th}"

                 f" ({resized_size_kb:.1f}KB)")

        return (rel_path, f"{orig_w}x{orig_h}",

                f"{orig_size_kb:.2f}", f"{resized_size_kb:.2f}",

                "Success", "")

    except MemoryError:

        return (rel_path, "", "", "0.00", "Failed", "Out of memory")

    except Exception as e:

        log.error(f"FAILED: {abs_path} - {e}\n{traceback.format_exc()}")

        return (rel_path, "", "", "0.00", "Failed", str(e))

# ─────────────────────────────────────────────────────────────────────────────

#  Audit Report

# ─────────────────────────────────────────────────────────────────────────────

def generate_audit_report(audit_data, out_dir, tw, th, mode, category):

    headers = ["File Name", "Original Dimensions",

               "Original Size (KB)", "Resized Size (KB)",

               "Status", "Failure Reason"]

    # Always write CSV

    csv_path = os.path.join(out_dir, "Resize_Audit_Report.csv")

    try:

        with open(csv_path, "w", newline="", encoding="utf-8") as f:

            w = csv.writer(f)

            w.writerow(headers)

            w.writerows(audit_data)

    except Exception as e:

        log.error(f"CSV write failed: {e}")

    if not HAS_OPENPYXL:

        return csv_path

    try:

        wb = openpyxl.Workbook()

        ws = wb.active

        ws.title = "Resize Audit"

        # Title row

        ws.merge_cells("A1:F1")

        t = ws["A1"]

        t.value = "IMAGE RESIZER BY SIAR DIGITAL — Audit Report"

        t.font = Font(name="Consolas", size=15, bold=True, color="00F5A0")

        t.fill = PatternFill("solid", fgColor="060910")

        t.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[1].height = 36

        # Info row

        ws.merge_cells("A2:F2")

        i = ws["A2"]

        i.value = (f"Created by Asif Nawaz  |  Siar Digital 2026  |  "

                   f"{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "

                   f"Target: {tw}×{th}  |  Mode: {mode}  |  Category: {category}")

        i.font = Font(name="Consolas", size=9, color="7A95C0")

        i.fill = PatternFill("solid", fgColor="0D1321")

        i.alignment = Alignment(horizontal="center")

        ws.row_dimensions[2].height = 20

        bdr = Border(

            left=Side(style="thin", color="1E2D45"),

            right=Side(style="thin", color="1E2D45"),

            top=Side(style="thin", color="1E2D45"),

            bottom=Side(style="thin", color="1E2D45"),

        )

        # Header row

        for ci, h in enumerate(headers, 1):

            c = ws.cell(row=3, column=ci, value=h)

            c.font = Font(name="Consolas", size=10, bold=True, color="E8F0FF")

            c.fill = PatternFill("solid", fgColor="1A2540")

            c.alignment = Alignment(horizontal="center")

            c.border = bdr

        sf = PatternFill("solid", fgColor="091A12")

        ff = PatternFill("solid", fgColor="1A0910")

        for ri, rd in enumerate(audit_data, 4):

            fail = "Failed" in str(rd[4])

            for ci, v in enumerate(rd, 1):

                c = ws.cell(row=ri, column=ci, value=v)

                c.border = bdr

                c.alignment = Alignment(horizontal="left")

                c.fill = ff if fail else sf

                c.font = Font(name="Consolas", size=9,

                              color="FF4D6D" if fail else "00F5A0")

        # Summary row

        sr = len(audit_data) + 5

        sc = sum(1 for r in audit_data if r[4] == "Success")

        fc = len(audit_data) - sc

        ws.merge_cells(f"A{sr}:F{sr}")

        s = ws.cell(row=sr, column=1,

                    value=f"TOTAL: {len(audit_data)}  |  "

                          f"SUCCESS: {sc}  |  FAILED: {fc}")

        s.font = Font(name="Consolas", size=11, bold=True, color="FFD166")

        s.fill = PatternFill("solid", fgColor="0D1321")

        s.alignment = Alignment(horizontal="center")

        for col in ws.columns:

            ml = max(len(str(c.value or "")) for c in col)

            ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 60)

        rp = os.path.join(out_dir, "Resize_Audit_Report.xlsx")

        wb.save(rp)

        log.info(f"Audit Excel: {rp}")

        return rp

    except Exception as e:

        log.error(f"Excel report failed: {e}")

        return csv_path

# ─────────────────────────────────────────────────────────────────────────────

#  Background Worker Thread

# ─────────────────────────────────────────────────────────────────────────────

def worker(input_path, tw, th, out_dir, turbo, mode, fmt,

           quality, max_threads, stop_evt, lossless,

           category, remove_watermark):

    try:

        input_path = os.path.normpath(os.path.abspath(input_path))

        out_dir = os.path.normpath(os.path.abspath(out_dir))

        file_list = scan_image_files(input_path)

        total = len(file_list)

        if total == 0:

            msg_queue.put(("no_files",))

            return

        msg_queue.put(("start", total))

        audit_data = []

        ok_count = fail_count = 0

        start_time = time.time()

        def _handle_result(res):

            nonlocal ok_count, fail_count

            audit_data.append(res)

            if res[4] == "Success":

                ok_count += 1

            else:

                fail_count += 1

        if turbo:

            with ThreadPoolExecutor(max_workers=max_threads) as pool:

                futures = {

                    pool.submit(

                        process_one, ap, rp, out_dir, tw, th, mode,

                        fmt, quality, lossless, False,

                        category, remove_watermark

                    ): rp

                    for ap, rp in file_list

                }

                for idx, future in enumerate(as_completed(futures)):

                    if stop_evt.is_set():

                        pool.shutdown(wait=False, cancel_futures=True)

                        break

                    try:

                        res = future.result(timeout=120)

                    except Exception as e:

                        res = (futures[future], "", "", "0.00",

                               "Failed", str(e))

                    _handle_result(res)

                    elapsed = time.time() - start_time

                    speed = (idx + 1) / max(elapsed, 0.001)

                    msg_queue.put(("progress", idx + 1, total,

                                   ok_count, fail_count, speed))

        else:

            for idx, (ap, rp) in enumerate(file_list):

                if stop_evt.is_set():

                    break

                res = process_one(ap, rp, out_dir, tw, th, mode,

                                  fmt, quality, lossless, True,

                                  category, remove_watermark)

                _handle_result(res)

                elapsed = time.time() - start_time

                speed = (idx + 1) / max(elapsed, 0.001)

                msg_queue.put(("progress", idx + 1, total,

                               ok_count, fail_count, speed))

        msg_queue.put(("generating_report",))

        report_path = generate_audit_report(

            audit_data, out_dir, tw, th, mode, category)

        elapsed_total = time.time() - start_time

        msg_queue.put(("done", ok_count, fail_count,

                       elapsed_total, report_path))

    except Exception as e:

        log.critical(f"WORKER CRASH: {e}\n{traceback.format_exc()}")

        msg_queue.put(("worker_error", str(e)))

# ─────────────────────────────────────────────────────────────────────────────

#  Tooltip

# ─────────────────────────────────────────────────────────────────────────────

class Tooltip:

    def __init__(self, widget, text):

        self.widget = widget

        self.text = text

        self.tipwin = None

        widget.bind("<Enter>", self.show)

        widget.bind("<Leave>", self.hide)

    def show(self, _=None):

        x = self.widget.winfo_rootx() + 20

        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6

        self.tipwin = tw = tk.Toplevel(self.widget)

        tw.wm_overrideredirect(True)

        tw.wm_geometry(f"+{x}+{y}")

        tw.configure(bg=C["surface3"])

        tk.Label(tw, text=self.text, bg=C["surface3"], fg=C["accent2"],

                 font=("Consolas", 9), relief="flat", bd=0,

                 padx=10, pady=5, wraplength=340, justify="left").pack()

    def hide(self, _=None):

        if self.tipwin:

            self.tipwin.destroy()

            self.tipwin = None

# ─────────────────────────────────────────────────────────────────────────────

#  Stat Card Widget

# ─────────────────────────────────────────────────────────────────────────────

class StatCard(ctk.CTkFrame):

    def __init__(self, master, label, color, **kw):

        super().__init__(master, fg_color=C["surface2"],

                         corner_radius=10, border_width=1,

                         border_color=color, **kw)

        self.val = ctk.StringVar(value="0")

        ctk.CTkLabel(self, text=label, font=FONTS["small"],

                     text_color=C["text2"]).pack(pady=(6, 0))

        ctk.CTkLabel(self, textvariable=self.val, font=FONTS["counter"],

                     text_color=color).pack(pady=(0, 6))

    def set(self, v):

        self.val.set(str(v))

# ─────────────────────────────────────────────────────────────────────────────

#  Main Application

# ─────────────────────────────────────────────────────────────────────────────

class ImageResizerPro(ctk.CTk):

    def __init__(self):

        super().__init__()

        ctk.set_appearance_mode("dark")

        ctk.set_default_color_theme("green")

        self.title("Image Resizer By Siar Digital")

        self.geometry("1280x960")

        self.minsize(1100, 820)

        self.configure(fg_color=C["bg"])

        self._stop_event = threading.Event()

        self._is_running = False

        self._tk_images = []

        # State vars

        self.folder_path   = ctk.StringVar()

        self.width_var     = ctk.StringVar(value="1080")

        self.height_var    = ctk.StringVar(value="1080")

        self.mode_var      = ctk.StringVar(value="Smart Crop (AI)")

        self.format_var    = ctk.StringVar(value="WEBP")

        self.quality_var   = ctk.IntVar(value=88)

        self.turbo_var     = ctk.BooleanVar(value=False)

        self.threads_var   = ctk.IntVar(value=4)

        self.preset_var    = ctk.StringVar(value="Instagram Post (1:1)")

        self.lossless_var  = ctk.BooleanVar(value=False)

        self.category_var  = ctk.StringVar(value="General")

        self.watermark_var = ctk.BooleanVar(value=False)

        self.status_var    = ctk.StringVar(value="Ready — Load a folder")

        self.quality_lbl   = ctk.StringVar(value="88%")

        self.thread_lbl    = ctk.StringVar(value="4 Threads")

        self.speed_var     = ctk.StringVar(value="0.0 img/s")

        self._build_ui()

        self._poll_queue()

        log.info("Siar Digital Image Resizer v4.0 started")

    # ── UI Builder ────────────────────────────────────────────────────────────

    def _build_ui(self):

        self._build_header()

        self._build_bottom_bar()

        self._build_log_strip()

        body = ctk.CTkFrame(self, fg_color="transparent")

        body.pack(fill="both", expand=True, padx=12, pady=(4, 0))

        left = ctk.CTkScrollableFrame(body, fg_color="transparent", width=560)

        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        right = ctk.CTkFrame(body, fg_color="transparent", width=420)

        right.pack(side="right", fill="both", padx=(6, 0))

        self._build_left(left)

        self._build_right(right)

    def _build_header(self):

        hdr = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=78)

        hdr.pack(fill="x")

        hdr.pack_propagate(False)

        left_hdr = ctk.CTkFrame(hdr, fg_color="transparent")

        left_hdr.pack(side="left", padx=20, pady=8)

        # Neon title

        ctk.CTkLabel(left_hdr,

                     text="▶  IMAGE RESIZER BY SIAR DIGITAL",

                     font=FONTS["hero"], text_color=C["accent"]).pack(anchor="w")

        ctk.CTkLabel(left_hdr,

                     text="Advanced Batch Engine v4.0  |  AI Subject Detection  |  "

                          "Quality Guard  |  Smart Fill",

                     font=FONTS["subtitle"], text_color=C["text2"]).pack(anchor="w")

        right_hdr = ctk.CTkFrame(hdr, fg_color="transparent")

        right_hdr.pack(side="right", padx=20)

        # Library badge row

        badge_row = ctk.CTkFrame(right_hdr, fg_color="transparent")

        badge_row.pack()

        for lib, has, col in [

            ("YOLO", HAS_YOLO, C["accent"]),

            ("OpenCV", HAS_OPENCV, C["accent2"]),

            ("rembg", HAS_REMBG, C["accent5"]),

            ("MediaPipe", HAS_MEDIAPIPE, C["accent6"]),

        ]:

            color = col if has else C["text3"]

            ctk.CTkLabel(badge_row, text=f"{'●' if has else '○'} {lib}",

                         font=FONTS["badge"], text_color=color).pack(

                side="left", padx=5)

        self.status_pill = ctk.CTkLabel(

            right_hdr, textvariable=self.status_var,

            font=FONTS["small"], text_color=C["accent2"],

            fg_color=C["surface2"], corner_radius=20, padx=14, pady=6)

        self.status_pill.pack(pady=(6, 0))

    def _build_log_strip(self):

        ctk.CTkLabel(self, text=f"Session log: {log_file}",

                     font=FONTS["tiny"], text_color=C["text3"]).pack(

            anchor="w", padx=18, pady=(2, 0))

    def _section(self, parent, title, color=None):

        color = color or C["accent"]

        f = ctk.CTkFrame(parent, fg_color=C["surface"],

                         corner_radius=12, border_width=1,

                         border_color=C["border"])

        f.pack(fill="x", pady=5)

        h = ctk.CTkFrame(f, fg_color=C["surface2"],

                         corner_radius=0, height=30)

        h.pack(fill="x")

        h.pack_propagate(False)

        ctk.CTkLabel(h, text=f"  {title}", font=FONTS["label"],

                     text_color=color).pack(side="left", padx=8)

        return f

    def _build_left(self, parent):

        # ── SOURCE FOLDER ─────────────────────────────────────────────────────

        sec = self._section(parent, "SOURCE FOLDER", C["accent2"])

        inner = ctk.CTkFrame(sec, fg_color="transparent")

        inner.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(inner,

                     text="Input folder → Output: {Name}_resized  (sub-folder tree mirrored)",

                     font=FONTS["small"], text_color=C["text2"]).pack(anchor="w")

        row = ctk.CTkFrame(inner, fg_color="transparent")

        row.pack(fill="x", pady=4)

        ctk.CTkEntry(row, textvariable=self.folder_path, height=36,

                     fg_color=C["surface2"], border_color=C["border"],

                     text_color=C["text"], font=FONTS["body"],

                     placeholder_text="Select input folder...").pack(

            side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(row, text="Browse", width=110, height=36,

                      fg_color=C["accent2"], hover_color="#00A8CC",

                      text_color="#000", font=FONTS["body"],

                      command=self._browse_input).pack(side="right")

        # ── SIZE & PRESETS ────────────────────────────────────────────────────

        sec2 = self._section(parent, "SIZE & PRESETS", C["accent"])

        inner2 = ctk.CTkFrame(sec2, fg_color="transparent")

        inner2.pack(fill="x", padx=12, pady=8)

        ctk.CTkOptionMenu(inner2, variable=self.preset_var,

                          values=list(PRESETS.keys()),

                          fg_color=C["surface2"], button_color=C["accent"],

                          button_hover_color="#00CC80", text_color=C["text"],

                          font=FONTS["body"],

                          dropdown_fg_color=C["surface"],

                          dropdown_text_color=C["text"],

                          command=self._apply_preset, height=36).pack(

            fill="x", pady=4)

        dr = ctk.CTkFrame(inner2, fg_color="transparent")

        dr.pack(fill="x")

        for lbl, var in [("Width (px)", self.width_var),

                         ("Height (px)", self.height_var)]:

            b = ctk.CTkFrame(dr, fg_color="transparent")

            b.pack(side="left", expand=True, fill="x", padx=3)

            ctk.CTkLabel(b, text=lbl, font=FONTS["small"],

                         text_color=C["text2"]).pack(anchor="w")

            ctk.CTkEntry(b, textvariable=var, height=40,

                         font=("Consolas", 16, "bold"),

                         text_color=C["accent"], fg_color=C["surface2"],

                         border_color=C["accent"],

                         justify="center").pack(fill="x")

        # ── CATEGORY MODE ─────────────────────────────────────────────────────

        sec_cat = self._section(parent, "CATEGORY MODE", C["accent6"])

        inner_cat = ctk.CTkFrame(sec_cat, fg_color="transparent")

        inner_cat.pack(fill="x", padx=12, pady=8)

        cat_row = ctk.CTkFrame(inner_cat, fg_color="transparent")

        cat_row.pack(fill="x")

        for cat, info in CATEGORIES.items():

            b = ctk.CTkFrame(cat_row, fg_color=C["surface2"], corner_radius=8)

            b.pack(side="left", expand=True, fill="x", padx=2)

            ctk.CTkRadioButton(b, text=f"{info['icon']}\n{cat}",

                               variable=self.category_var, value=cat,

                               fg_color=C["accent6"],

                               hover_color=C["accent4"],

                               text_color=C["text"],

                               font=FONTS["tiny"]).pack(pady=6, padx=4)

            Tooltip(b, info["desc"])

        # ── PLACEMENT MODE ────────────────────────────────────────────────────

        sec3 = self._section(parent, "PLACEMENT MODE", C["accent5"])

        inner3 = ctk.CTkFrame(sec3, fg_color="transparent")

        inner3.pack(fill="x", padx=12, pady=8)

        for mn, tip in MODE_TIPS.items():

            is_ai = "AI" in mn

            bg_color = "#1A1530" if is_ai else C["surface2"]

            bc = C["accent5"] if is_ai else C["border"]

            r = ctk.CTkFrame(inner3, fg_color=bg_color, corner_radius=8,

                             border_width=1, border_color=bc)

            r.pack(fill="x", pady=2)

            ctk.CTkRadioButton(r, text=mn, variable=self.mode_var, value=mn,

                               fg_color=C["accent5"] if is_ai else C["accent"],

                               hover_color=C["accent5"],

                               text_color=C["text"],

                               font=FONTS["body"]).pack(side="left",

                                                        padx=10, pady=6)

            if is_ai:

                badge = "AI POWERED" if (HAS_YOLO or HAS_OPENCV) else "needs opencv"

                ctk.CTkLabel(r, text=f"[{badge}]",

                             font=FONTS["badge"],

                             text_color=C["accent5"]).pack(side="left", padx=4)

            Tooltip(r, tip)

        # ── FORMAT & QUALITY ──────────────────────────────────────────────────

        sec4 = self._section(parent, "FORMAT & QUALITY", C["accent4"])

        inner4 = ctk.CTkFrame(sec4, fg_color="transparent")

        inner4.pack(fill="x", padx=12, pady=8)

        fr = ctk.CTkFrame(inner4, fg_color="transparent")

        fr.pack(fill="x")

        for fm in ["WEBP", "JPEG", "PNG", "TIFF", "BMP"]:

            b = ctk.CTkFrame(fr, fg_color=C["surface2"], corner_radius=8)

            b.pack(side="left", padx=2, expand=True, fill="x")

            ctk.CTkRadioButton(b, text=fm, variable=self.format_var, value=fm,

                               fg_color=C["accent4"],

                               hover_color=C["accent4"],

                               text_color=C["text"],

                               font=FONTS["small"]).pack(pady=6, padx=4)

        qr = ctk.CTkFrame(inner4, fg_color="transparent")

        qr.pack(fill="x", pady=(8, 0))

        ctk.CTkLabel(qr, text="Quality:", font=FONTS["small"],

                     text_color=C["text2"]).pack(side="left")

        ctk.CTkLabel(qr, text="Quality Guard: ≤1.5 MB auto-optimize",

                     font=FONTS["tiny"],

                     text_color=C["accent7"]).pack(side="right")

        ctk.CTkLabel(inner4, textvariable=self.quality_lbl,

                     font=FONTS["label"],

                     text_color=C["accent4"]).pack(anchor="e")

        ctk.CTkSlider(inner4, from_=10, to=100, variable=self.quality_var,

                      button_color=C["accent4"],

                      progress_color=C["accent4"],

                      command=lambda v: self.quality_lbl.set(f"{int(v)}%"),

                      height=14).pack(fill="x", pady=3)

        # ── ADVANCED ──────────────────────────────────────────────────────────

        sec5 = self._section(parent, "ADVANCED OPTIONS", C["accent3"])

        inner5 = ctk.CTkFrame(sec5, fg_color="transparent")

        inner5.pack(fill="x", padx=12, pady=8)

        for sw_text, sw_var, sw_col, tip_txt in [

            ("Zero-Loss Mode (Lossless PNG, ignores quality slider)",

             self.lossless_var, C["accent"],

             "Forces PNG lossless output. File size not compressed."),

            ("TURBO Mode (Multi-thread batch, disables live preview)",

             self.turbo_var, C["accent2"],

             "Parallel processing. Fastest for large batches. No preview."),

            ("AI Watermark / Logo Removal (rembg inpainting)" +

             ("" if HAS_REMBG else "  ← pip install rembg"),

             self.watermark_var, C["accent5"],

             "Uses rembg U2-Net to remove background logos/watermarks."),

        ]:

            f = ctk.CTkFrame(inner5, fg_color=C["surface2"], corner_radius=8)

            f.pack(fill="x", pady=3)

            sw = ctk.CTkSwitch(f, text=sw_text, variable=sw_var,

                               progress_color=sw_col, font=FONTS["small"],

                               text_color=C["text"],

                               command=self._toggle_lossless

                               if sw_var == self.lossless_var else None)

            sw.pack(side="left", padx=10, pady=8)

            Tooltip(f, tip_txt)

        tr = ctk.CTkFrame(inner5, fg_color="transparent")

        tr.pack(fill="x", pady=(6, 0))

        ctk.CTkLabel(tr, text="Turbo Threads:", font=FONTS["small"],

                     text_color=C["text2"]).pack(side="left")

        ctk.CTkLabel(tr, textvariable=self.thread_lbl,

                     font=FONTS["label"],

                     text_color=C["accent"]).pack(side="right")

        ctk.CTkSlider(inner5, from_=1, to=16, number_of_steps=15,

                      variable=self.threads_var,

                      button_color=C["accent"],

                      progress_color=C["accent"],

                      command=lambda v: self.thread_lbl.set(

                          f"{int(v)} Threads"),

                      height=14).pack(fill="x", pady=3)

    def _build_right(self, parent):

        # ── Live Preview ──────────────────────────────────────────────────────

        ps = ctk.CTkFrame(parent, fg_color=C["surface"],

                          corner_radius=12, border_width=1,

                          border_color=C["accent5"])

        ps.pack(fill="x", pady=(0, 6))

        ph = ctk.CTkFrame(ps, fg_color=C["surface2"],

                          corner_radius=0, height=30)

        ph.pack(fill="x")

        ph.pack_propagate(False)

        ctk.CTkLabel(ph, text="  ◉  LIVE PREVIEW MONITOR",

                     font=FONTS["label"],

                     text_color=C["accent5"]).pack(side="left", padx=8)

        self.prev_filename = ctk.CTkLabel(

            ps, text="Waiting for processing...",

            font=FONTS["small"], text_color=C["text3"])

        self.prev_filename.pack(pady=(8, 3))

        ir = ctk.CTkFrame(ps, fg_color="transparent")

        ir.pack(fill="x", padx=8, pady=(2, 8))

        ob = ctk.CTkFrame(ir, fg_color=C["surface2"],

                          corner_radius=8, border_width=1,

                          border_color=C["accent2"])

        ob.pack(side="left", expand=True, fill="both", padx=(0, 3))

        ctk.CTkLabel(ob, text="ORIGINAL", font=FONTS["small"],

                     text_color=C["accent2"]).pack(pady=(6, 2))

        self.lbl_orig = ctk.CTkLabel(ob, text="No Preview",

                                     text_color=C["text3"],

                                     font=FONTS["small"],

                                     width=165, height=150)

        self.lbl_orig.pack(padx=4, pady=(0, 8))

        rb = ctk.CTkFrame(ir, fg_color=C["surface2"],

                          corner_radius=8, border_width=1,

                          border_color=C["accent"])

        rb.pack(side="left", expand=True, fill="both", padx=(3, 0))

        ctk.CTkLabel(rb, text="RESIZED", font=FONTS["small"],

                     text_color=C["accent"]).pack(pady=(6, 2))

        self.lbl_res = ctk.CTkLabel(rb, text="No Preview",

                                    text_color=C["text3"],

                                    font=FONTS["small"],

                                    width=165, height=150)

        self.lbl_res.pack(padx=4, pady=(0, 8))

        # ── Statistics ────────────────────────────────────────────────────────

        ss = ctk.CTkFrame(parent, fg_color=C["surface"],

                          corner_radius=12, border_width=1,

                          border_color=C["border"])

        ss.pack(fill="x", pady=6)

        sh = ctk.CTkFrame(ss, fg_color=C["surface2"],

                          corner_radius=0, height=30)

        sh.pack(fill="x")

        sh.pack_propagate(False)

        ctk.CTkLabel(sh, text="  STATISTICS", font=FONTS["label"],

                     text_color=C["accent4"]).pack(side="left", padx=8)

        cr = ctk.CTkFrame(ss, fg_color="transparent")

        cr.pack(fill="x", padx=8, pady=8)

        self.card_total  = StatCard(cr, "TOTAL",  C["accent2"])

        self.card_done   = StatCard(cr, "DONE",   C["accent"])

        self.card_failed = StatCard(cr, "FAILED", C["accent3"])

        for c in [self.card_total, self.card_done, self.card_failed]:

            c.pack(side="left", expand=True, fill="both", padx=3)

        ctk.CTkLabel(ss, textvariable=self.speed_var,

                     font=FONTS["label"],

                     text_color=C["accent"]).pack(pady=(0, 8))

        # ── Progress ──────────────────────────────────────────────────────────

        ps2 = ctk.CTkFrame(parent, fg_color=C["surface"],

                           corner_radius=12, border_width=1,

                           border_color=C["border"])

        ps2.pack(fill="x", pady=6)

        ctk.CTkLabel(ps2, text="  PROGRESS", font=FONTS["label"],

                     text_color=C["text2"]).pack(anchor="w", padx=10,

                                                 pady=(6, 2))

        self.prog_bar = ctk.CTkProgressBar(ps2, progress_color=C["accent"],

                                           height=16, corner_radius=8)

        self.prog_bar.set(0)

        self.prog_bar.pack(fill="x", padx=10, pady=(0, 4))

        self.prog_label = ctk.CTkLabel(ps2, text="0 / 0 images",

                                       font=FONTS["small"],

                                       text_color=C["text2"])

        self.prog_label.pack(pady=(0, 6))

        # ── Lib Status ────────────────────────────────────────────────────────

        ls = ctk.CTkFrame(parent, fg_color=C["surface"],

                          corner_radius=12, border_width=1,

                          border_color=C["border"])

        ls.pack(fill="x", pady=4)

        ctk.CTkLabel(ls, text="  AI LIBRARY STATUS",

                     font=FONTS["label"],

                     text_color=C["text2"]).pack(anchor="w", padx=10, pady=(8,4))

        for lib, has, install_cmd, desc in [

            ("YOLOv8 (Ultralytics)", HAS_YOLO,

             "pip install ultralytics", "Deep object detection"),

            ("OpenCV", HAS_OPENCV,

             "pip install opencv-python", "Core CV & face detection"),

            ("rembg (U2-Net)", HAS_REMBG,

             "pip install rembg", "Watermark/BG removal"),

            ("MediaPipe", HAS_MEDIAPIPE,

             "pip install mediapipe", "Segmentation guard"),

            ("openpyxl", HAS_OPENPYXL,

             "pip install openpyxl", "Excel audit reports"),

        ]:

            row = ctk.CTkFrame(ls, fg_color="transparent")

            row.pack(fill="x", padx=10, pady=1)

            dot = "✓" if has else "✗"

            col = C["success"] if has else C["error"]

            ctk.CTkLabel(row, text=f"{dot} {lib}",

                         font=FONTS["small"], text_color=col,

                         width=180, anchor="w").pack(side="left")

            if not has:

                ctk.CTkLabel(row, text=install_cmd,

                             font=FONTS["tiny"],

                             text_color=C["text3"]).pack(side="left", padx=6)

        ctk.CTkFrame(ls, height=6, fg_color="transparent").pack()

        # ── Creator Credit ────────────────────────────────────────────────────

        cf = ctk.CTkFrame(parent, fg_color=C["surface"],

                          corner_radius=12, border_width=2,

                          border_color=C["gold"])

        cf.pack(fill="x", pady=(8, 4))

        ctk.CTkLabel(cf, text="CREATED BY ASIF NAWAZ",

                     font=FONTS["credit"],

                     text_color=C["gold"]).pack(pady=(12, 2))

        ctk.CTkLabel(cf, text="Siar Digital 2026  |  All Rights Reserved",

                     font=FONTS["credit_sm"],

                     text_color=C["accent2"]).pack(pady=(0, 2))

        ctk.CTkLabel(cf, text="www.siardigital.com",

                     font=FONTS["credit_sm"],

                     text_color=C["text2"]).pack(pady=(0, 12))

    def _build_bottom_bar(self):

        self.bottom_bar = ctk.CTkFrame(self, fg_color=C["surface"],

                                       corner_radius=0, height=78)

        self.bottom_bar.pack(fill="x", side="bottom")

        self.bottom_bar.pack_propagate(False)

        btn_row = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")

        btn_row.pack(expand=True, fill="both", padx=18, pady=12)

        self.start_btn = ctk.CTkButton(

            btn_row, text="▶  START PROCESSING",

            fg_color=C["accent"], hover_color="#00CC80",

            text_color="#000000", height=54,

            font=FONTS["big_btn"], corner_radius=12,

            command=self._start)

        self.start_btn.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.stop_btn = ctk.CTkButton(

            btn_row, text="■  STOP",

            fg_color=C["accent3"], hover_color="#CC2244",

            text_color="#FFF", height=54, width=130,

            font=FONTS["big_btn"], corner_radius=12,

            command=self._stop, state="disabled")

        self.stop_btn.pack(side="left", padx=(0, 8))

        self.open_btn = ctk.CTkButton(

            btn_row, text="📂  Open Output",

            fg_color=C["surface2"], hover_color=C["border2"],

            text_color=C["text2"], height=54, width=160,

            font=FONTS["body"], corner_radius=12,

            border_width=1, border_color=C["border2"],

            command=self._open_output)

        self.open_btn.pack(side="left")

    # ── Actions ───────────────────────────────────────────────────────────────

    def _browse_input(self):

        p = filedialog.askdirectory(title="Select Input Folder")

        if p:

            p = os.path.normpath(os.path.abspath(p))

            self.folder_path.set(p)

            fl = scan_image_files(p)

            self.card_total.set(len(fl))

            self._set_status(

                f"Loaded: {os.path.basename(p)} ({len(fl)} images)",

                C["accent2"])

            if len(fl) == 0:

                messagebox.showwarning(

                    "No Images",

                    f"No supported images found in:\n{p}")

    def _apply_preset(self, name):

        dims = PRESETS.get(name)

        if dims:

            self.width_var.set(str(dims[0]))

            self.height_var.set(str(dims[1]))

    def _toggle_lossless(self):

        if self.lossless_var.get():

            self.format_var.set("PNG")

            self.quality_var.set(100)

            self.quality_lbl.set("100% (Lossless)")

        else:

            self.quality_lbl.set(f"{self.quality_var.get()}%")

    def _set_status(self, text, color=None):

        self.status_var.set(text)

        if color:

            self.status_pill.configure(text_color=color)

    def _start(self):

        inp = self.folder_path.get().strip()

        if not inp or not os.path.isdir(inp):

            messagebox.showwarning("Invalid Input",

                                   "Please select a valid input folder.")

            return

        try:

            tw = int(self.width_var.get())

            th = int(self.height_var.get())

            assert 1 <= tw <= 20000 and 1 <= th <= 20000

        except Exception:

            messagebox.showerror("Invalid Size",

                                 "Width/Height must be between 1 and 20000.")

            return

        inp = os.path.normpath(os.path.abspath(inp))

        out_dir = os.path.join(os.path.dirname(inp),

                               f"{os.path.basename(inp)}_resized")

        try:

            os.makedirs(out_dir, exist_ok=True)

        except Exception as e:

            messagebox.showerror("Output Error", str(e))

            return

        self._stop_event.clear()

        self._is_running = True

        self.prog_bar.set(0)

        self.prog_label.configure(text="Scanning...")

        self.card_done.set(0)

        self.card_failed.set(0)

        self.speed_var.set("0.0 img/s")

        self.start_btn.configure(state="disabled")

        self.stop_btn.configure(state="normal")

        self._set_status("Processing...", C["accent4"])

        self.prev_filename.configure(text="Processing...",

                                     text_color=C["accent4"])

        self.lbl_orig.configure(image=None, text="Processing...")

        self.lbl_res.configure(image=None, text="Processing...")

        threading.Thread(

            target=worker, daemon=True,

            args=(inp, tw, th, out_dir,

                  self.turbo_var.get(),

                  self.mode_var.get(),

                  self.format_var.get(),

                  self.quality_var.get(),

                  int(self.threads_var.get()),

                  self._stop_event,

                  self.lossless_var.get(),

                  self.category_var.get(),

                  self.watermark_var.get())

        ).start()

        log.info(f"Job: {inp} → {out_dir} [{tw}×{th}]  "

                 f"mode={self.mode_var.get()}  "

                 f"cat={self.category_var.get()}")

    def _stop(self):

        self._stop_event.set()

        self._set_status("Stopping...", C["accent4"])

        self.stop_btn.configure(state="disabled")

    def _open_output(self):

        inp = self.folder_path.get().strip()

        if not inp:

            messagebox.showinfo("Open Output", "No folder loaded.")

            return

        inp = os.path.normpath(os.path.abspath(inp))

        path = os.path.join(os.path.dirname(inp),

                            f"{os.path.basename(inp)}_resized")

        if not os.path.isdir(path):

            messagebox.showinfo("Open Output",

                                "Output folder not found.\nProcess images first.")

            return

        try:

            if sys.platform == "win32":

                os.startfile(path)

            elif sys.platform == "darwin":

                os.system(f'open "{path}"')

            else:

                os.system(f'xdg-open "{path}"')

        except Exception as e:

            messagebox.showerror("Error", str(e))

    # ── Queue Polling ─────────────────────────────────────────────────────────

    def _poll_queue(self):

        try:

            while True:

                msg = msg_queue.get_nowait()

                kind = msg[0]

                if kind == "start":

                    self.card_total.set(msg[1])

                    self.prog_label.configure(text=f"0 / {msg[1]} images")

                elif kind == "progress":

                    _, done, total, ok, fail, speed = msg

                    frac = done / max(total, 1)

                    self.prog_bar.set(frac)

                    self.prog_label.configure(

                        text=f"{done} / {total}  ({frac*100:.1f}%)")

                    self.card_done.set(ok)

                    self.card_failed.set(fail)

                    self.speed_var.set(f"{speed:.1f} img/s")

                elif kind == "preview":

                    _, p_orig, p_res, fname, fsize = msg

                    self.prev_filename.configure(

                        text=f"{fname}  ({fsize:.1f} KB)",

                        text_color=C["accent2"])

                    tk_o = ImageTk.PhotoImage(p_orig)

                    self.lbl_orig.configure(image=tk_o, text="")

                    self.lbl_orig.image = tk_o

                    tk_r = ImageTk.PhotoImage(p_res)

                    self.lbl_res.configure(image=tk_r, text="")

                    self.lbl_res.image = tk_r

                    self._tk_images.extend([tk_o, tk_r])

                    if len(self._tk_images) > 24:

                        self._tk_images = self._tk_images[-24:]

                elif kind == "generating_report":

                    self._set_status("Generating Audit Report...",

                                     C["accent5"])

                elif kind == "done":

                    _, ok, fail, elapsed, report_path = msg

                    self._is_running = False

                    self.start_btn.configure(state="normal")

                    self.stop_btn.configure(state="disabled")

                    self.prog_bar.set(1.0)

                    self._set_status(

                        f"Done: {ok} processed, {fail} failed", C["accent"])

                    rpt = (f"\n\nAudit Report:\n{report_path}"

                           if report_path else "")

                    messagebox.showinfo(

                        "Processing Complete — Siar Digital",

                        f"IMAGE RESIZER BY SIAR DIGITAL\n"

                        f"CREATED BY ASIF NAWAZ\n"

                        f"{'─'*40}\n\n"

                        f"  Processed : {ok}\n"

                        f"  Failed    : {fail}\n"

                        f"  Time      : {elapsed:.1f}s\n"

                        f"  Speed     : {ok/max(elapsed,0.001):.1f} img/s\n"

                        f"\nLog: {log_file}{rpt}")

                elif kind == "no_files":

                    self._is_running = False

                    self.start_btn.configure(state="normal")

                    self.stop_btn.configure(state="disabled")

                    self._set_status("No images found", C["accent3"])

                    messagebox.showwarning("No Images",

                                          "No supported images found in that folder.")

                elif kind == "worker_error":

                    self._is_running = False

                    self.start_btn.configure(state="normal")

                    self.stop_btn.configure(state="disabled")

                    self._set_status("Error!", C["accent3"])

                    messagebox.showerror(

                        "Internal Error",

                        f"Worker crashed:\n{msg[1]}\n\nLog: {log_file}")

        except queue.Empty:

            pass

        except Exception as e:

            log.error(f"Poll error: {e}")

        self.after(40, self._poll_queue)

    def on_closing(self):

        if self._is_running:

            if messagebox.askyesno("Quit",

                                   "Processing is running. Stop and quit?"):

                self._stop_event.set()

                self.destroy()

        else:

            self.destroy()

# ─────────────────────────────────────────────────────────────────────────────

#  Entry Point

# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    try:

        app = ImageResizerPro()

        app.protocol("WM_DELETE_WINDOW", app.on_closing)

        app.mainloop()

    except Exception as e:

        log.critical(f"FATAL: {e}\n{traceback.format_exc()}")

        try:

            messagebox.showerror(

                "Fatal Error",

                f"Image Resizer By Siar Digital crashed:\n\n{e}\n\nLog: {log_file}")

        except Exception:

            pass

        sys.exit(1)
