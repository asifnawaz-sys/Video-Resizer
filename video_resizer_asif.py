"""

╔══════════════════════════════════════════════════════════════════════════════╗
║   VIDEO RESIZER by ASIF NAWAZ  --  SIAR DIGITAL  v5.2                        ║
║   Auto MP4 fallback | Perfect smart crop | YOLOv8 | Shopify guard           ║
╚══════════════════════════════════════════════════════════════════════════════╝

DEPENDENCIES:
    pip install moviepy opencv-python Pillow ttkbootstrap ultralytics

RUN:
    python video_resizer_asif.py

"""

import os
import sys
import json
import shutil
import logging
import zipfile
import threading
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

# ---------- optional imports ----------
LIB_STATUS = {}

def _try_import(name, pip_name=None):
    try:
        mod = __import__(name)
        ver = getattr(mod, "__version__", "?")
        LIB_STATUS[name] = {"ok": True, "version": ver, "error": ""}
        return mod
    except Exception as exc:
        LIB_STATUS[name] = {"ok": False, "version": "", "error": str(exc)}
        return None

cv2_mod      = _try_import("cv2", "opencv-python")
PIL_mod      = _try_import("PIL", "Pillow")
tb_mod       = _try_import("ttkbootstrap", "ttkbootstrap")
ultra_mod    = _try_import("ultralytics", "ultralytics")
moviepy_mod  = _try_import("moviepy", "moviepy")

HAS_CV2     = LIB_STATUS["cv2"]["ok"]
HAS_PIL     = LIB_STATUS["PIL"]["ok"]
HAS_TB      = LIB_STATUS["ttkbootstrap"]["ok"]
HAS_YOLO    = LIB_STATUS["ultralytics"]["ok"]
HAS_MOVIEPY = LIB_STATUS["moviepy"]["ok"]

if HAS_CV2:
    import cv2
if HAS_PIL:
    from PIL import Image, ImageTk
if HAS_TB:
    import ttkbootstrap as tb
if HAS_YOLO:
    from ultralytics import YOLO as _YOLO

# ---------- logging ----------
LOG_DIR  = Path(os.path.expanduser("~")) / "SiarResizerLogs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("SiarResizer")
log.info("Session started. Log: %s", LOG_FILE)

# ---------- constants ----------
VIDEO_EXTS = frozenset({".mp4", ".mov", ".avi", ".mkv", ".webm",
                        ".wmv", ".flv", ".m4v", ".3gp", ".ts", ".mts"})

PRESETS = {
    "Custom":                                  (None, None),
    "Shopify Product (4:5)  850 x 1280":       (850,  1280),
    "Instagram Reel (9:16)  1080 x 1920":      (1080, 1920),
    "TikTok (9:16)  1080 x 1920":              (1080, 1920),
    "YouTube Shorts (9:16)  1080 x 1920":      (1080, 1920),
    "Instagram Post (1:1)  1080 x 1080":       (1080, 1080),
    "Instagram Landscape (4:5)  1080 x 1350":  (1080, 1350),
    "Facebook Story (9:16)  1080 x 1920":      (1080, 1920),
    "YouTube (16:9)  1920 x 1080":             (1920, 1080),
    "Twitter/X Landscape  1280 x 720":         (1280, 720),
    "LinkedIn (1:1)  1080 x 1080":             (1080, 1080),
    "Pinterest (2:3)  1000 x 1500":            (1000, 1500),
    "Amazon Product  1200 x 1200":             (1200, 1200),
    "4K UHD  3840 x 2160":                     (3840, 2160),
    "Full HD  1920 x 1080":                    (1920, 1080),
    "HD  1280 x 720":                          (1280, 720),
    "SD  854 x 480":                           (854,  480),
}

YOLO_CLASSES = [
    "person","bicycle","car","motorcycle","airplane","bus","train","truck",
    "boat","traffic light","fire hydrant","stop sign","parking meter","bench",
    "bird","cat","dog","horse","sheep","cow","elephant","bear","zebra","giraffe",
    "backpack","umbrella","handbag","tie","suitcase","frisbee","skis",
    "snowboard","sports ball","kite","baseball bat","baseball glove",
    "skateboard","surfboard","tennis racket","bottle","wine glass","cup",
    "fork","knife","spoon","bowl","banana","apple","sandwich","orange",
    "broccoli","carrot","hot dog","pizza","donut","cake","chair","couch",
    "potted plant","bed","dining table","toilet","tv","laptop","mouse",
    "remote","keyboard","cell phone","microwave","oven","toaster","sink",
    "refrigerator","book","clock","vase","scissors","teddy bear",
    "hair drier","toothbrush",
]

CATEGORY_MODES = [
    ("General",   "#00ff88"),
    ("Clothing",  "#00cfff"),
    ("Jewelry",   "#f59e0b"),
    ("Furniture", "#a78bfa"),
    ("Portrait",  "#fb7185"),
    ("Product",   "#34d399"),
]

PLACEMENT_MODES = [
    ("Smart Crop  (Center)",      "smart_crop",     False),
    ("YOLOv8 Subject Crop",       "yolo_crop",      True ),
    ("Letterbox  (Blur BG)",      "letterbox_blur", False),
    ("Letterbox  (Dark BG)",      "letterbox_dark", False),
    ("Letterbox  (White BG)",     "letterbox_white",False),
    ("Mirror BG  (True Mirror)",  "mirror_bg",      False),
    ("AI Background Extend",      "ai_bg_extend",   True ),
    ("Fill & Crop  (Center)",     "fill_crop",      False),
    ("Stretch to Fit",            "stretch",        False),
    ("Pad Custom  (Black BG)",    "pad_custom",     False),
]

OUTPUT_FORMATS = ["MP4", "MOV", "MKV", "WEBM"]
CODEC_MAP = {
    "MP4":  ("libx264",    "aac",     ".mp4"),
    "MOV":  ("libx264",    "aac",     ".mov"),
    "MKV":  ("libx264",    "aac",     ".mkv"),
    "WEBM": ("libvpx-vp9", "libopus", ".webm"),
}
QUALITY_PRESETS = {"Maximum  (CRF 18)": "18", "High     (CRF 22)": "22",
                   "Balanced (CRF 26)": "26", "Smaller  (CRF 30)": "30", "Tiny     (CRF 35)": "35"}
SPEED_PRESETS  = ["ultrafast","superfast","veryfast","faster","fast","medium","slow","slower","veryslow"]
FPS_OPTIONS    = ["Source FPS","23.976","24","25","29.97","30","50","60"]
AUDIO_BITRATES = ["96k","128k","192k","256k","320k","No Audio"]

# ---------- colour palette ----------
class P:
    BG0  = "#020509"; BG1  = "#050b14"; BG2  = "#080f1c"; BG3  = "#0c1522"
    BG4  = "#101c2c"; BG5  = "#141f30"; BDR  = "#1a2d45"; BDR2 = "#243d5e"
    G1 = "#00ff88"; G2 = "#00cc66"; B1 = "#00d4ff"; B2 = "#009fcc"
    PU = "#9d6cff"; PU2= "#6d28d9"; OR = "#ff9f00"; RD = "#ff4050"
    RD2= "#cc2233"; PK = "#f472b6"; YL = "#facc15"; CY = "#22d3ee"
    TXT  = "#ccddef"; TXT2 = "#5a7a9a"; TXT3 = "#2d4460"
    MONO   = ("Consolas", 9); MONO_S = ("Consolas", 7); MONO_B = ("Consolas", 9, "bold")
    MONO_L = ("Consolas", 12, "bold"); MONO_H = ("Consolas", 16, "bold")
    DET_COLORS = [G1,B1,PU,OR,RD,PK,YL,"#34d399","#60a5fa","#c084fc","#fb923c"]

# ---------- ffmpeg helpers ----------
def _find_ffmpeg():
    """Find ffmpeg - check all possible bundled locations, then PATH, then Homebrew"""
    if getattr(sys, 'frozen', False):
        search_dirs = []
        if hasattr(sys, '_MEIPASS'):
            search_dirs.append(Path(sys._MEIPASS))
        exe_dir = Path(sys.executable).parent
        search_dirs.append(exe_dir)
        search_dirs.append(exe_dir / '_internal')
        search_dirs.append(exe_dir.parent / 'Frameworks')
        for search_dir in search_dirs:
            ffmpeg_p = search_dir / 'ffmpeg'
            if ffmpeg_p.exists():
                os.environ['PATH'] = str(search_dir) + os.pathsep + os.environ.get('PATH', '')
                break
    ffmpeg = shutil.which('ffmpeg')
    ffprobe = shutil.which('ffprobe')
    if not ffmpeg:
        for p in ['/usr/local/bin/ffmpeg', '/opt/homebrew/bin/ffmpeg', '/usr/bin/ffmpeg']:
            if Path(p).exists():
                ffmpeg = p
                break
    if not ffprobe:
        for p in ['/usr/local/bin/ffprobe', '/opt/homebrew/bin/ffprobe', '/usr/bin/ffprobe']:
            if Path(p).exists():
                ffprobe = p
                break
    return ffmpeg, ffprobe

FFMPEG_PATH, FFPROBE_PATH = _find_ffmpeg()

def _check_ffmpeg():
    if not FFMPEG_PATH:
        return False, ""
    try:
        r = subprocess.run([FFMPEG_PATH, "-version"], capture_output=True, text=True, timeout=10)
        if r.returncode == 0:
            ver = r.stdout.splitlines()[0] if r.stdout else "unknown"
            return True, ver
        return False, ""
    except:
        return False, ""

HAS_FFMPEG, FFMPEG_VERSION = _check_ffmpeg()

_ffmpeg_encoders_cache = None
def get_available_encoders():
    global _ffmpeg_encoders_cache
    if _ffmpeg_encoders_cache is not None:
        return _ffmpeg_encoders_cache
    encoders = {"libx264": False, "libvpx-vp9": False, "aac": False, "libopus": False}
    if not HAS_FFMPEG:
        return encoders
    try:
        r = subprocess.run([FFMPEG_PATH, "-encoders"], capture_output=True, text=True, timeout=10)
        if r.returncode == 0:
            for line in r.stdout.splitlines():
                if "libx264" in line: encoders["libx264"] = True
                if "libvpx-vp9" in line: encoders["libvpx-vp9"] = True
                if "aac" in line and "EA" in line: encoders["aac"] = True
                if "libopus" in line: encoders["libopus"] = True
    except Exception as e:
        log.warning("Failed to probe encoders: %s", e)
    _ffmpeg_encoders_cache = encoders
    return encoders

# ---------- YOLO engine ----------
_yolo_model = None
_yolo_model_lock = threading.Lock()

def get_yolo():
    global _yolo_model
    if not HAS_YOLO:
        return None
    with _yolo_model_lock:
        if _yolo_model is None:
            try:
                _yolo_model = _YOLO("yolov8n.pt")
                log.info("YOLOv8n model loaded.")
            except Exception as exc:
                log.error("YOLOv8 load failed: %s", exc)
                _yolo_model = None
    return _yolo_model

def yolo_detect(frame_bgr, tw, th):
    try:
        model = get_yolo()
        if model is None:
            return None, []
        results = model(frame_bgr, verbose=False)[0]
        if results.boxes is None or len(results.boxes) == 0:
            return None, []
        boxes = results.boxes.xyxy.cpu().numpy()
        confs = results.boxes.conf.cpu().numpy()
        cls_ids = results.boxes.cls.cpu().numpy().astype(int)
        dets = []
        for i, (box, conf, cid) in enumerate(zip(boxes, confs, cls_ids)):
            name = YOLO_CLASSES[cid] if cid < len(YOLO_CLASSES) else f"cls_{cid}"
            color = P.DET_COLORS[i % len(P.DET_COLORS)]
            dets.append({"class": name, "conf": float(conf), "bbox": tuple(int(v) for v in box), "color": color})
        dets.sort(key=lambda d: d["conf"], reverse=True)
        if not dets:
            return None, []
        h, w = frame_bgr.shape[:2]
        bx1, by1, bx2, by2 = dets[0]["bbox"]
        cx = (bx1 + bx2) / 2.0
        cy = (by1 + by2) / 2.0
        tr = tw / th
        sr = w / h
        if sr > tr:
            cw = int(h * tr)
            x1 = int(cx - cw / 2)
            x1 = max(0, min(x1, w - cw))
            crop = (x1, 0, x1 + cw, h)
        else:
            ch = int(w / tr)
            y1 = int(cy - ch / 2)
            y1 = max(0, min(y1, h - ch))
            crop = (0, y1, w, y1 + ch)
        return crop, dets
    except Exception as exc:
        log.warning("YOLO detection error: %s", exc)
        return None, []

def draw_yolo_overlay(frame_bgr, detections, max_draw=10):
    if not HAS_CV2 or not detections:
        return frame_bgr
    try:
        out = frame_bgr.copy()
        for det in detections[:max_draw]:
            x1, y1, x2, y2 = det["bbox"]
            r,g,b = _hex2bgr(det["color"])
            cv2.rectangle(out, (x1,y1),(x2,y2),(b,g,r), 2)
            lbl = f"{det['class']} {det['conf']:.0%}"
            (tw2,th2),_ = cv2.getTextSize(lbl,cv2.FONT_HERSHEY_SIMPLEX,0.45,1)
            cv2.rectangle(out,(x1,y1-th2-8),(x1+tw2+6,y1),(b,g,r),-1)
            cv2.putText(out,lbl,(x1+3,y1-3), cv2.FONT_HERSHEY_SIMPLEX,0.45,(0,0,0),1,cv2.LINE_AA)
        return out
    except Exception:
        return frame_bgr

def _hex2bgr(h):
    try:
        h = h.lstrip("#")
        r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return r,g,b
    except:
        return 255,255,255

def scan_folder_for_videos(root_folder):
    results = []
    root = Path(root_folder).resolve()
    for dirpath, _, filenames in os.walk(root):
        for fname in filenames:
            if Path(fname).suffix.lower() in VIDEO_EXTS:
                abs_path = Path(dirpath) / fname
                try:
                    rel_path = abs_path.relative_to(root)
                    results.append({"abs": str(abs_path), "rel": str(rel_path)})
                except:
                    pass
    return results

def _ffprobe_json(path, select_streams="v:0", show_entries="stream=width,height,r_frame_rate"):
    probe = FFPROBE_PATH or "ffprobe"
    try:
        cmd = [probe, "-v", "quiet", "-print_format", "json",
               "-select_streams", select_streams, "-show_entries", show_entries, str(path)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if r.returncode == 0 and r.stdout.strip():
            return json.loads(r.stdout)
    except:
        pass
    return {}

def _get_video_info(path):
    info = {"width": 1920, "height": 1080, "fps": 30.0, "duration": 30.0, "has_audio": True}
    try:
        jv = _ffprobe_json(path, select_streams="v:0",
                           show_entries="stream=width,height,r_frame_rate,duration")
        streams = jv.get("streams", [])
        if streams:
            s = streams[0]
            if s.get("width"):  info["width"]  = int(s["width"])
            if s.get("height"): info["height"] = int(s["height"])
            rfr = s.get("r_frame_rate", "30/1")
            try:
                num, den = rfr.split("/")
                info["fps"] = round(float(num) / max(float(den), 1), 3)
            except:
                pass
            if s.get("duration"):
                info["duration"] = float(s["duration"])
        jf = _ffprobe_json(path, select_streams="", show_entries="format=duration")
        d = jf.get("format", {}).get("duration")
        if d:
            info["duration"] = float(d)
        ja = _ffprobe_json(path, select_streams="a:0", show_entries="stream=codec_type")
        info["has_audio"] = bool(ja.get("streams"))
        if HAS_CV2 and (info["width"] == 0 or info["height"] == 0):
            cap = cv2.VideoCapture(str(path))
            if cap.isOpened():
                info["width"]  = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))  or 1920
                info["height"] = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT)) or 1080
                fps_cv = cap.get(cv2.CAP_PROP_FPS)
                if fps_cv and fps_cv > 0: info["fps"] = fps_cv
                dur_cv = cap.get(cv2.CAP_PROP_FRAME_COUNT) / max(fps_cv, 1)
                if dur_cv > 0: info["duration"] = dur_cv
                cap.release()
    except Exception as exc:
        log.warning("_get_video_info error: %s", exc)
    info["width"]    = max(info["width"], 1)
    info["height"]   = max(info["height"], 1)
    info["fps"]      = max(info["fps"], 1.0)
    info["duration"] = max(info["duration"], 0.1)
    return info

def _read_mid_frame(path, vi=None):
    if not HAS_CV2:
        return None
    try:
        cap = cv2.VideoCapture(str(path))
        if not cap.isOpened():
            return None
        total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        cap.set(cv2.CAP_PROP_POS_FRAMES, max(0, total // 2))
        ret, frame = cap.read()
        cap.release()
        return frame if ret else None
    except:
        return None

def build_vf(placement, w_in, h_in, tw, th, crop_bbox=None):
    tw_even = tw if (tw % 2 == 0) else tw + 1
    th_even = th if (th % 2 == 0) else th + 1
    if tw_even != tw or th_even != th:
        log.warning(f"Target dimensions {tw}x{th} not even -> using {tw_even}x{th_even}")
    try:
        def _cover_crop():
            return (f"scale={tw_even}:{th_even}:force_original_aspect_ratio=increase:flags=lanczos,"
                    f"crop={tw_even}:{th_even}:trunc((iw-{tw_even})/2):trunc((ih-{th_even})/2)")
        def _letterbox(color):
            return (f"scale={tw_even}:{th_even}:force_original_aspect_ratio=decrease:flags=lanczos,"
                    f"pad={tw_even}:{th_even}:(ow-iw)/2:(oh-ih)/2:color={color}")
        def _blur_bg():
            return (f"split[fg][bg];[bg]scale={tw_even}:{th_even}:force_original_aspect_ratio=increase:flags=lanczos,"
                    f"crop={tw_even}:{th_even},gblur=sigma=30[blurbg];"
                    f"[fg]scale={tw_even}:{th_even}:force_original_aspect_ratio=decrease:flags=lanczos[fgs];"
                    f"[blurbg][fgs]overlay=(W-w)/2:(H-h)/2,format=yuv420p")
        def _mirror_bg():
            return (f"[0:v]scale={tw_even}:{th_even}:force_original_aspect_ratio=increase:flags=lanczos,"
                    f"crop={tw_even}:{th_even}[base];[base]hflip[base_hflip];[base]vflip[base_vflip];"
                    f"[base_vflip]hflip[base_hvflip];[base][base_hflip]hstack=inputs=2[top];"
                    f"[base_vflip][base_hvflip]hstack=inputs=2[bottom];[top][bottom]vstack=inputs=2[mirror_tile];"
                    f"[mirror_tile]scale={tw_even}:{th_even}:flags=lanczos[background];"
                    f"[0:v]scale={tw_even}:{th_even}:force_original_aspect_ratio=decrease:flags=lanczos[fg];"
                    f"[background][fg]overlay=(W-w)/2:(H-h)/2,format=yuv420p")
        if placement in ("smart_crop", "fill_crop"):
            return _cover_crop()
        if placement == "yolo_crop":
            if crop_bbox:
                x1, y1, x2, y2 = crop_bbox
                cw = max(x2 - x1, 2); ch = max(y2 - y1, 2)
                cw = cw if (cw % 2 == 0) else cw + 1
                ch = ch if (ch % 2 == 0) else ch + 1
                return f"crop={cw}:{ch}:{x1}:{y1},scale={tw_even}:{th_even}:flags=lanczos"
            return _cover_crop()
        if placement == "stretch":
            return f"scale={tw_even}:{th_even}:flags=lanczos"
        if placement in ("letterbox_dark", "pad_custom"):
            return _letterbox("black")
        if placement == "letterbox_white":
            return _letterbox("white")
        if placement in ("letterbox_blur", "ai_bg_extend"):
            return _blur_bg()
        if placement == "mirror_bg":
            return _mirror_bg()
        return _cover_crop()
    except Exception as exc:
        log.error("build_vf error: %s", exc)
        return f"scale={tw_even}:{th_even}:flags=lanczos"

def _target_video_kbps(duration_sec, max_bytes, audio_kbps):
    CONTAINER_OVERHEAD = 256 * 1024
    audio_bytes = (audio_kbps * 1000 / 8) * duration_sec
    video_budget = max(max_bytes - CONTAINER_OVERHEAD - audio_bytes, 256 * 1024)
    kbps = (video_budget * 8) / (duration_sec * 1000)
    return max(int(kbps), 80)

def _run_ffmpeg(cmd, label="ffmpeg"):
    ff = FFMPEG_PATH or "ffmpeg"
    cmd = [ff if str(c) == "ffmpeg" else str(c) for c in cmd]
    log.debug("Running: %s", " ".join(str(c) for c in cmd))
    try:
        r = subprocess.run([str(c) for c in cmd], capture_output=True, text=True, timeout=3600)
        if r.returncode != 0:
            log.warning("%s failed (rc=%d):\n%s", label, r.returncode, r.stderr[-1200:])
            return False, r.stderr
        return True, r.stderr
    except Exception as exc:
        log.error("%s exception: %s", label, exc)
        return False, str(exc)

def encode_file(in_path, out_path, cfg, vf, vi, log_cb):
    fmt      = cfg["format"]
    crf      = cfg["crf"]
    speed    = cfg["speed"]
    fps_opt  = cfg["fps"]
    audio_br = cfg["audio_bitrate"]
    max_bytes= cfg["max_bytes"]
    codec, acodec, ext = CODEC_MAP[fmt]
    duration   = vi["duration"]
    has_audio  = vi["has_audio"]
    no_audio   = (audio_br == "No Audio") or (not has_audio)
    audio_kbps = 0 if no_audio else int(audio_br.replace("k", ""))
    is_complex = any(x in vf for x in ("split", "overlay", "hstack", "vstack"))

    def build_cmd(vcodec, acodec_choice, audio_present, extra_flags=None):
        cmd = ["ffmpeg", "-y", "-i", str(in_path)]
        if fps_opt != "Source FPS":
            cmd += ["-r", fps_opt]
        if is_complex:
            cmd += ["-filter_complex", vf]
        else:
            cmd += ["-vf", vf]
        cmd += ["-c:v", vcodec]
        if fmt == "WEBM":
            cmd += ["-crf", crf, "-b:v", "0", "-deadline", "good", "-cpu-used", "2"]
        else:
            cmd += ["-crf", crf, "-preset", speed, "-profile:v", "high", "-level:v", "4.2"]
        cmd += ["-pix_fmt", "yuv420p", "-movflags", "+faststart"]
        if audio_present and not no_audio:
            cmd += ["-c:a", acodec_choice, "-b:a", audio_br, "-ar", "44100", "-ac", "2"]
        else:
            cmd += ["-an"]
        if extra_flags:
            cmd += extra_flags
        cmd += [str(out_path)]
        return cmd

    def try_encode(vcodec, acodec_choice, audio_present, label_suffix=""):
        cmd = build_cmd(vcodec, acodec_choice, audio_present)
        label = f"CRF-pass{label_suffix}"
        ok, stderr = _run_ffmpeg(cmd, label)
        if ok:
            return True, stderr
        if audio_present and not no_audio and ("encoder" in stderr.lower() and ("aac" in stderr.lower() or "opus" in stderr.lower())):
            log_cb(f"   {label}: audio encoder failed - retrying without audio", "warn")
            cmd_no_audio = build_cmd(vcodec, acodec_choice, False)
            ok2, stderr2 = _run_ffmpeg(cmd_no_audio, label + " (no audio)")
            if ok2:
                return True, stderr2
        return False, stderr

    log_cb(f"   CRF={crf}  speed={speed}  format={fmt}", "info")
    ok, stderr = try_encode(codec, acodec, True)
    if not ok:
        if fmt == "WEBM" and "libopus" in stderr:
            log_cb("   WEBM fallback to MP4", "warn")
            new_fmt = "MP4"
            new_codec, new_acodec, new_ext = CODEC_MAP[new_fmt]
            new_out_path = Path(str(out_path)).with_suffix(new_ext)
            cmd_fallback = ["ffmpeg", "-y", "-i", str(in_path)]
            if fps_opt != "Source FPS": cmd_fallback += ["-r", fps_opt]
            if is_complex: cmd_fallback += ["-filter_complex", vf]
            else: cmd_fallback += ["-vf", vf]
            cmd_fallback += ["-c:v", new_codec, "-crf", crf, "-preset", speed,
                             "-profile:v", "high", "-level:v", "4.2",
                             "-pix_fmt", "yuv420p", "-movflags", "+faststart"]
            if not no_audio: cmd_fallback += ["-c:a", new_acodec, "-b:a", audio_br, "-ar", "44100", "-ac", "2"]
            else: cmd_fallback += ["-an"]
            cmd_fallback += [str(new_out_path)]
            ok2, stderr2 = _run_ffmpeg(cmd_fallback, "CRF-fallback-MP4")
            if ok2:
                out_path = new_out_path; fmt = new_fmt; codec = new_codec
                acodec = new_acodec; ext = new_ext; ok = True; cfg["format"] = new_fmt
        if not ok:
            log_cb("   CRF pass failed - retrying ultrafast...", "warn")
            cmd_uf = build_cmd(codec, acodec, True, extra_flags=["-preset", "ultrafast"])
            ok_uf, stderr_uf = _run_ffmpeg(cmd_uf, "CRF-ultrafast")
            if ok_uf:
                ok = True
            else:
                cmd_uf_no_audio = build_cmd(codec, acodec, False, extra_flags=["-preset", "ultrafast"])
                ok_uf2, _ = _run_ffmpeg(cmd_uf_no_audio, "CRF-ultrafast-no-audio")
                ok = ok_uf2
            if not ok:
                log_cb("   ENCODE FAILED - check log file.", "error")
                return False, 0.0

    try:
        actual = Path(str(out_path)).stat().st_size
    except:
        actual = 0
    actual_mb = actual / (1024*1024)
    log_cb(f"   Output: {actual_mb:.2f} MB", "info")

    if max_bytes < 999 * 1024 * 1024 and actual > max_bytes:
        vid_kbps = _target_video_kbps(duration, max_bytes, audio_kbps)
        log_cb(f"   Over limit - size-guard 2-pass -> {vid_kbps} kbps", "warn")
        def _bitrate_flags(kbps, pass_num=1):
            cmd = ["-c:v", codec, "-b:v", f"{kbps}k"]
            if fmt == "WEBM": cmd += ["-deadline", "good", "-cpu-used", "2"]
            else: cmd += ["-preset", speed, "-profile:v", "high", "-level:v", "4.2"]
            if pass_num == 1:
                null = "/dev/null" if sys.platform != "win32" else "NUL"
                cmd += ["-pass", "1", "-f", "mp4" if fmt != "WEBM" else "webm", null]
            else:
                cmd += ["-pass", "2"]
            cmd += ["-pix_fmt", "yuv420p", "-movflags", "+faststart"]
            return cmd
        pass1_cmd = ["ffmpeg", "-y", "-i", str(in_path)]
        if fps_opt != "Source FPS": pass1_cmd += ["-r", fps_opt]
        if is_complex: pass1_cmd += ["-filter_complex", vf]
        else: pass1_cmd += ["-vf", vf]
        pass1_cmd += _bitrate_flags(vid_kbps, pass_num=1)
        if not no_audio: pass1_cmd += ["-c:a", acodec, "-b:a", audio_br, "-ar", "44100", "-ac", "2"]
        else: pass1_cmd += ["-an"]
        ok1, _ = _run_ffmpeg(pass1_cmd, "2pass-1st")
        if ok1:
            pass2_cmd = ["ffmpeg", "-y", "-i", str(in_path)]
            if fps_opt != "Source FPS": pass2_cmd += ["-r", fps_opt]
            if is_complex: pass2_cmd += ["-filter_complex", vf]
            else: pass2_cmd += ["-vf", vf]
            pass2_cmd += _bitrate_flags(vid_kbps, pass_num=2)
            if not no_audio: pass2_cmd += ["-c:a", acodec, "-b:a", audio_br, "-ar", "44100", "-ac", "2"]
            else: pass2_cmd += ["-an"]
            pass2_cmd += [str(out_path)]
            ok2, stderr2 = _run_ffmpeg(pass2_cmd, "2pass-2nd")
            if ok2:
                try: actual = Path(str(out_path)).stat().st_size
                except: pass
                actual_mb = actual / (1024*1024)
                log_cb(f"   After 2-pass: {actual_mb:.2f} MB", "ok")
        for f in Path(str(out_path)).parent.glob("ffmpeg2pass-0*"):
            try: f.unlink()
            except: pass

    return True, actual_mb

class VideoProcessor:
    def __init__(self, log_cb, progress_cb, preview_cb, detection_cb):
        self.log_cb = log_cb
        self.progress_cb = progress_cb
        self.preview_cb = preview_cb
        self.detection_cb = detection_cb
        self.stop_flag = threading.Event()

    def stop(self):
        self.stop_flag.set()

    def _log(self, msg, level="info"):
        try:
            self.log_cb(msg, level)
            getattr(log, level if level in ("debug","info","warning","error","critical") else "info", log.info)(msg)
        except:
            pass

    def process_batch(self, file_list, cfg):
        self.stop_flag.clear()
        total = len(file_list)
        done = 0
        failed = 0
        out_root = Path(cfg["output_dir"])
        try:
            out_root.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self._log(f"Cannot create output folder: {exc}", "error")
            return
        self._log(f"Batch start -- {total} file(s)", "head")
        self._log(f"Output -> {out_root}", "info")
        for idx, item in enumerate(file_list):
            if self.stop_flag.is_set():
                self._log("Stopped by user.", "warn")
                break
            in_path = Path(item["abs"])
            rel_path = Path(item["rel"])
            fmt_ext = CODEC_MAP[cfg["format"]][2]
            out_name = rel_path.stem + fmt_ext
            out_rel = rel_path.parent / out_name
            out_path = out_root / out_rel
            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                self._log(f"mkdir error ({out_path.parent}): {exc}", "error")
                failed += 1
                self.progress_cb(idx+1, total, done, failed)
                continue
            if cfg["skip_done"] and out_path.exists() and out_path.stat().st_size > 0:
                self._log(f"SKIP (exists): {out_rel}", "ok")
                done += 1
                self.progress_cb(idx+1, total, done, failed)
                continue
            self._log(f"FILE [{idx+1}/{total}]  {rel_path.name}", "head")
            if not in_path.exists():
                self._log(f"  Input not found: {in_path}", "error")
                failed += 1
                self.progress_cb(idx+1, total, done, failed)
                continue
            if in_path.stat().st_size == 0:
                self._log(f"  Input is empty: {in_path}", "error")
                failed += 1
                self.progress_cb(idx+1, total, done, failed)
                continue
            try:
                file_cfg = cfg.copy()
                success, mb = self._process_one(in_path, out_path, file_cfg)
                if success:
                    done += 1
                    self._log(f"DONE: {out_name}  ({mb:.2f} MB)", "ok")
                    self.preview_cb(str(out_path))
                else:
                    failed += 1
                    self._log(f"FAILED: {out_name}", "error")
                    try:
                        if out_path.exists(): out_path.unlink()
                    except: pass
            except Exception as exc:
                failed += 1
                self._log(f"EXCEPTION [{rel_path.name}]: {exc}", "error")
                log.error("Unhandled exception:\n%s", traceback.format_exc())
            self.progress_cb(idx+1, total, done, failed)
        if cfg.get("auto_zip") and done > 0:
            self._make_zip(str(out_root))
        self._log(f"BATCH COMPLETE  --  {done} OK  |  {failed} FAILED  |  {total} TOTAL", "head")

    def _process_one(self, in_path, out_path, cfg):
        tw = int(cfg["width"])
        th = int(cfg["height"])
        placement = cfg["placement"]
        vi = _get_video_info(str(in_path))
        self._log(f"   Probe: {vi['width']}x{vi['height']}  {vi['fps']:.2f}fps  {vi['duration']:.1f}s  audio={'yes' if vi['has_audio'] else 'no'}", "info")
        self._log(f"   Target: {tw}x{th}  [{placement}]", "info")
        crop_bbox = None
        detections = []
        if placement == "yolo_crop":
            if HAS_YOLO and HAS_CV2:
                self._log("   YOLOv8: scanning frame...", "yolo")
                frame = _read_mid_frame(str(in_path), vi)
                if frame is not None:
                    crop_bbox, detections = yolo_detect(frame, tw, th)
                    if detections:
                        top = ", ".join(f"{d['class']} {d['conf']:.0%}" for d in detections[:5])
                        self._log(f"   Detected: {top}", "yolo")
                        try: self.detection_cb(frame, detections)
                        except: pass
                    else:
                        self._log("   No subject - using smart crop", "warn")
                        placement = "smart_crop"
                else:
                    self._log("   Frame read failed - using smart crop", "warn")
                    placement = "smart_crop"
            else:
                self._log("   YOLOv8/OpenCV unavailable - using smart crop", "warn")
                placement = "smart_crop"
        vf = build_vf(placement, vi["width"], vi["height"], tw, th, crop_bbox)
        self._log(f"   Filter: {vf[:120]}{'...' if len(vf)>120 else ''}", "info")
        return encode_file(str(in_path), str(out_path), cfg, vf, vi, self._log)

    def _make_zip(self, folder):
        zip_path = folder.rstrip("/\\") + "_Siar_Resized.zip"
        self._log(f"Creating ZIP: {Path(zip_path).name}", "info")
        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
                for dirpath, _, files in os.walk(folder):
                    for f in files:
                        full = Path(dirpath) / f
                        arc = full.relative_to(folder)
                        zf.write(str(full), str(arc))
            self._log(f"ZIP ready: {Path(zip_path).name}", "ok")
        except Exception as exc:
            self._log(f"ZIP error: {exc}", "error")

class BarMeter(tk.Canvas):
    def __init__(self, parent, color=P.G1, height=5, **kw):
        super().__init__(parent, height=height, bg=P.BG0, highlightthickness=0, **kw)
        self._color = color
        self._val = 0.0
        self.bind("<Configure>", lambda e: self._draw())
    def set_value(self, v):
        self._val = max(0.0, min(1.0, float(v)))
        self._draw()
    def _draw(self):
        try:
            self.delete("all")
            w = self.winfo_width()
            h = self.winfo_height()
            if w < 2: return
            self.create_rectangle(0,0,w,h,fill=P.BG0,outline="")
            filled = int(w * self._val)
            if filled > 0:
                self.create_rectangle(0,0,filled,h,fill=self._color,outline="")
        except:
            pass

class NeonBtn(tk.Button):
    def __init__(self, parent, text, command=None, color=P.G1, danger=False, **kw):
        c = P.RD if danger else color
        super().__init__(parent, text=text, command=command,
                         font=("Consolas",8,"bold"), fg=c, bg=P.BG4,
                         activebackground=P.BDR2, activeforeground="#fff",
                         relief="flat", cursor="hand2", padx=10, pady=5,
                         highlightthickness=1, highlightbackground=c, **kw)
        self.bind("<Enter>", lambda e: self.config(bg=P.BDR2, fg="#fff"))
        self.bind("<Leave>", lambda e: self.config(bg=P.BG4, fg=c))

class DetectionPanel(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=P.BG2, **kw)
        self._build()
    def _build(self):
        hdr = tk.Frame(self, bg=P.BG2); hdr.pack(fill="x", padx=8, pady=(6,3))
        tk.Label(hdr, text="YOLOv8  DETECTION  MONITOR", font=P.MONO_B, fg=P.PU, bg=P.BG2).pack(side="left")
        col = P.G1 if HAS_YOLO else P.RD
        txt = "LOADED" if HAS_YOLO else "NOT INSTALLED"
        tk.Label(hdr, text=txt, font=P.MONO_S, fg=col, bg=P.BG2).pack(side="right")
        self.img_cv = tk.Canvas(self, bg=P.BG0, height=190, highlightthickness=1, highlightbackground=P.BDR)
        self.img_cv.pack(fill="x", padx=8, pady=(0,4))
        self._no_det_text()
        tk.Label(self, text="  YOLOv8n  |  COCO 80 classes  |  Pre-trained", font=("Consolas",7), fg=P.TXT3, bg=P.BG3).pack(fill="x", padx=8, pady=(0,3))
        ch = tk.Frame(self, bg=P.BG2); ch.pack(fill="x", padx=10)
        for t,w in [("#",2),("CLASS",14),("CONF",6),("SCORE",0)]:
            tk.Label(ch, text=t, font=("Consolas",7,"bold"), fg=P.TXT3, bg=P.BG2, width=w, anchor="w").pack(side="left", padx=(0,3))
        lw = tk.Frame(self, bg=P.BG3, highlightthickness=1, highlightbackground=P.BDR); lw.pack(fill="x", padx=8, pady=(2,8))
        self._rows = []
        for i in range(10):
            col = P.DET_COLORS[i % len(P.DET_COLORS)]
            row = tk.Frame(lw, bg=P.BG3); row.pack(fill="x", padx=4, pady=1)
            tk.Label(row, text=f"{i+1}", font=P.MONO_S, fg=P.TXT3, bg=P.BG3, width=2).pack(side="left")
            tk.Label(row, text=" ", bg=col, width=2).pack(side="left", padx=(2,0))
            nm = tk.Label(row, text="", font=("Consolas",8), fg=P.TXT, bg=P.BG3, width=14, anchor="w"); nm.pack(side="left", padx=(4,0))
            cf = tk.Label(row, text="", font=("Consolas",8,"bold"), fg=col, bg=P.BG3, width=6); cf.pack(side="left")
            bar = BarMeter(row, color=col, height=5); bar.pack(side="left", fill="x", expand=True, padx=(4,4))
            self._rows.append({"name":nm,"conf":cf,"bar":bar})
    def _no_det_text(self):
        try:
            self.img_cv.delete("all")
            cw = max(self.img_cv.winfo_width(), 360)
            self.img_cv.create_text(cw//2, 95, text="No detection yet", fill=P.TXT3, font=P.MONO_S)
        except:
            pass
    def update_detection(self, frame_bgr, detections):
        if not HAS_CV2 or not HAS_PIL: return
        try:
            ann = draw_yolo_overlay(frame_bgr, detections)
            rgb = cv2.cvtColor(ann, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb)
            cw = max(self.img_cv.winfo_width(), 360)
            img.thumbnail((cw, 190))
            imgtk = ImageTk.PhotoImage(img)
            self.img_cv.delete("all")
            self.img_cv.config(height=max(img.height, 40))
            self.img_cv.create_image(cw//2, img.height//2, anchor="center", image=imgtk)
            self.img_cv._img = imgtk
        except:
            pass
        try:
            for i, rd in enumerate(self._rows):
                if i < len(detections):
                    d = detections[i]
                    rd["name"].config(text=d["class"])
                    rd["conf"].config(text=f"{d['conf']:.0%}")
                    rd["bar"].set_value(d["conf"])
                else:
                    rd["name"].config(text="")
                    rd["conf"].config(text="")
                    rd["bar"].set_value(0)
        except:
            pass
    def clear(self):
        try:
            self._no_det_text()
            for rd in self._rows:
                rd["name"].config(text="")
                rd["conf"].config(text="")
                rd["bar"].set_value(0)
        except:
            pass

class VideoResizerApp:
    _LOGO_FRAMES = [P.G1,"#00ee77","#00dd66",P.G1,P.B1,"#00bbdd",P.B1,P.G1]
    _logo_idx = 0

    def __init__(self, root):
        self.root = root
        self.root.title("Video Resizer by Asif Nawaz  |  SIAR DIGITAL  v5.2")
        self.root.configure(bg=P.BG1)
        self.root.geometry("1500x900")
        self.root.minsize(1300, 800)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.input_files = []
        self.processor = None
        self._proc_thread = None
        self._placement_frames = {}
        self._cat_frames = {}
        self._init_vars()
        self._apply_style()
        self._build_ui()
        self._on_preset_change(None)
        self._pulse_logo()
        if not HAS_FFMPEG:
            self.root.after(500, self._warn_ffmpeg)

    def _init_vars(self):
        self.preset_var    = tk.StringVar(value="Shopify Product (4:5)  850 x 1280")
        self.width_var     = tk.StringVar(value="850")
        self.height_var    = tk.StringVar(value="1280")
        self.category_var  = tk.StringVar(value="General")
        self.placement_var = tk.StringVar(value="smart_crop")
        self.format_var    = tk.StringVar(value="MP4")
        self.quality_var   = tk.StringVar(value="Balanced (CRF 26)")
        self.speed_var     = tk.StringVar(value="slow")
        self.fps_var       = tk.StringVar(value="Source FPS")
        self.audio_var     = tk.StringVar(value="192k")
        self.output_dir    = tk.StringVar(value=str(Path.home() / "Desktop" / "output_videos"))
        self.autozip_var   = tk.BooleanVar(value=True)
        self.skip_done_var = tk.BooleanVar(value=True)
        self.shopify_var   = tk.BooleanVar(value=True)
        self.max_mb_var    = tk.StringVar(value="7")

    def _apply_style(self):
        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TCombobox", fieldbackground=P.BG4, background=P.BG3,
                            foreground=P.TXT, selectbackground=P.PU2,
                            selectforeground="#fff", bordercolor=P.BDR2, arrowcolor=P.G1)
            style.map("TCombobox", fieldbackground=[("readonly", P.BG4)],
                      foreground=[("readonly", P.TXT)])
            style.configure("Green.Horizontal.TProgressbar", troughcolor=P.BG0,
                            background=P.G1, bordercolor=P.BG0, lightcolor=P.G1, darkcolor=P.G2)
        except:
            pass

    def _build_ui(self):
        self._build_topbar()
        body = tk.Frame(self.root, bg=P.BG1); body.pack(fill="both", expand=True)
        left_wrap = tk.Frame(body, bg=P.BG1); left_wrap.pack(side="left", fill="both", expand=True)
        self._build_scrollable_left(left_wrap)
        right = tk.Frame(body, bg=P.BG0, width=400); right.pack(side="right", fill="y"); right.pack_propagate(False)
        self._build_right_panel(right)
        self._build_bottombar()

    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=P.BG0, height=56); bar.pack(fill="x"); bar.pack_propagate(False)
        left = tk.Frame(bar, bg=P.BG0); left.pack(side="left", fill="y", padx=(10,0))
        self.logo_lbl = tk.Label(left, text="🎬  VIDEO RESIZER  🎬", font=P.MONO_H, fg=P.G1, bg=P.BG0)
        self.logo_lbl.pack(side="left", pady=8)
        tk.Label(left, text="  by Asif Nawaz  |  SIAR DIGITAL  ", font=("Consolas",8,"bold"), fg=P.YL, bg=P.BG0).pack(side="left", padx=10)
        tags = [("YOLOv8",P.PU,HAS_YOLO), ("COCO 80-class",P.B1,HAS_YOLO), ("Smart Crop",P.G1,True),
                ("Mirror BG",P.CY,True), ("Shopify Guard",P.OR,True), ("Folder Mirror",P.YL,True),
                ("CRF+2pass",P.PK,True), ("ffmpeg",P.G1,HAS_FFMPEG)]
        for tag,col,ok in tags:
            c = col if ok else P.TXT3
            f = tk.Frame(left, bg=P.BG3, highlightthickness=1, highlightbackground=c if ok else P.BDR)
            f.pack(side="left", padx=2, pady=15)
            tk.Label(f, text=f"  {tag}  ", font=("Consolas",7,"bold"), fg=c, bg=P.BG3).pack()
        pills = tk.Frame(bar, bg=P.BG0); pills.pack(side="right", padx=10)
        for name, ok in [("YOLOv8",HAS_YOLO),("OpenCV",HAS_CV2),("moviepy",HAS_MOVIEPY),
                         ("PIL",HAS_PIL),("ttk",HAS_TB),("ffmpeg",HAS_FFMPEG)]:
            col = P.G1 if ok else P.RD
            bg2 = "#041408" if ok else "#140408"
            f = tk.Frame(pills, bg=bg2, highlightthickness=1, highlightbackground=col)
            f.pack(side="left", padx=2, pady=14)
            tk.Label(f, text=f"  {'OK' if ok else 'NO'}  {name}  ", font=("Consolas",7,"bold"), fg=col, bg=bg2).pack()

    def _build_scrollable_left(self, parent):
        canvas = tk.Canvas(parent, bg=P.BG1, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.sf = tk.Frame(canvas, bg=P.BG1)
        self.sf.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.sf, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        def _scroll(event):
            try: canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        canvas.bind_all("<MouseWheel>", _scroll)
        col_a = tk.Frame(self.sf, bg=P.BG1); col_a.pack(side="left", fill="both", expand=True, padx=(6,3), pady=4)
        col_b = tk.Frame(self.sf, bg=P.BG1); col_b.pack(side="left", fill="both", expand=True, padx=(3,6), pady=4)
        self._build_input_panel(col_a)
        self._build_size_panel(col_a)
        self._build_category_panel(col_a)
        self._build_format_panel(col_a)
        self._build_advanced_panel(col_a)
        self._build_placement_panel(col_b)
        self._build_shopify_panel(col_b)
        self._build_output_panel(col_b)

    def _card(self, parent, title, accent=P.G1):
        outer = tk.Frame(parent, bg=P.BG1); outer.pack(fill="x", pady=(6,0))
        hdr = tk.Frame(outer, bg=P.BG2); hdr.pack(fill="x")
        tk.Frame(hdr, bg=accent, width=4).pack(side="left", fill="y")
        tk.Label(hdr, text=f"  {title}  ", font=P.MONO_B, fg=accent, bg=P.BG2, pady=6).pack(side="left")
        body = tk.Frame(outer, bg=P.BG3, highlightthickness=1, highlightbackground=P.BDR); body.pack(fill="x")
        return body

    def _build_input_panel(self, p):
        body = self._card(p, "INPUT FILES", P.G1)
        br = tk.Frame(body, bg=P.BG3); br.pack(fill="x", padx=10, pady=(10,4))
        NeonBtn(br, "Add Files",  self._add_files,   P.G1).pack(side="left", padx=(0,4))
        NeonBtn(br, "Add Folder", self._add_folder,  P.B1).pack(side="left", padx=(0,4))
        NeonBtn(br, "Clear All",  self._clear_files, danger=True).pack(side="right")
        self.file_count_lbl = tk.Label(body, text="No files queued", font=P.MONO_S, fg=P.TXT3, bg=P.BG3)
        self.file_count_lbl.pack(anchor="w", padx=12, pady=(0,4))
        lbf = tk.Frame(body, bg=P.BG0, highlightthickness=1, highlightbackground=P.BDR); lbf.pack(fill="x", padx=10, pady=(0,2))
        sb2 = tk.Scrollbar(lbf, bg=P.BG3); sb2.pack(side="right", fill="y")
        self.file_lb = tk.Listbox(lbf, bg=P.BG0, fg=P.TXT, font=("Consolas",7), height=7,
                                  selectbackground=P.PU2, borderwidth=0, highlightthickness=0, yscrollcommand=sb2.set)
        self.file_lb.pack(side="left", fill="x", expand=True)
        sb2.config(command=self.file_lb.yview)
        self.struct_lbl = tk.Label(body, text="", font=("Consolas",7), fg=P.OR, bg=P.BG3)
        self.struct_lbl.pack(anchor="w", padx=12, pady=(0,8))

    def _build_size_panel(self, p):
        body = self._card(p, "SIZE  &  PRESETS", P.B1)
        cb = ttk.Combobox(body, textvariable=self.preset_var, values=list(PRESETS.keys()), state="readonly", font=P.MONO)
        cb.pack(fill="x", padx=10, pady=(10,6))
        cb.bind("<<ComboboxSelected>>", self._on_preset_change)
        dr = tk.Frame(body, bg=P.BG3); dr.pack(fill="x", padx=10, pady=(0,10))
        for lbl, var, col in [("WIDTH  px", self.width_var, P.B1), ("HEIGHT px", self.height_var, P.PK)]:
            cf = tk.Frame(dr, bg=P.BG3); cf.pack(side="left", fill="x", expand=True, padx=(0,8))
            tk.Label(cf, text=lbl, font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(anchor="w")
            tk.Entry(cf, textvariable=var, font=("Consolas",18,"bold"), fg=col, bg=P.BG0, insertbackground=col,
                     relief="flat", highlightthickness=1, highlightbackground=P.BDR, highlightcolor=col).pack(fill="x", ipady=8)

    def _build_category_panel(self, p):
        body = self._card(p, "CATEGORY MODE", P.PU)
        row = tk.Frame(body, bg=P.BG3); row.pack(fill="x", padx=10, pady=8)
        for name, color in CATEGORY_MODES:
            f = tk.Frame(row, bg=P.BG4, cursor="hand2", highlightthickness=1, highlightbackground=P.BDR)
            f.pack(side="left", fill="x", expand=True, padx=2)
            rb = tk.Radiobutton(f, text=name, variable=self.category_var, value=name,
                                bg=P.BG4, fg=color, selectcolor=P.BG4, activebackground=P.BG4,
                                font=("Consolas",8,"bold"),
                                command=lambda c=color, fr=f, n=name: self._sel_cat(n, c))
            rb.pack(pady=8, padx=4)
            self._cat_frames[name] = (f, color)

    def _sel_cat(self, name, color):
        try:
            for n, (f, c) in self._cat_frames.items():
                f.configure(bg=P.BG4, highlightbackground=P.BDR)
                for w in f.winfo_children():
                    try: w.configure(bg=P.BG4)
                    except: pass
            fs, _ = self._cat_frames[name]
            fs.configure(bg=P.BG2, highlightbackground=color)
            for w in fs.winfo_children():
                try: w.configure(bg=P.BG2)
                except: pass
        except: pass

    def _build_format_panel(self, p):
        body = self._card(p, "FORMAT  &  QUALITY", P.OR)
        fr = tk.Frame(body, bg=P.BG3); fr.pack(fill="x", padx=10, pady=(10,4))
        tk.Label(fr, text="FORMAT:", font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(side="left", padx=(0,8))
        for fmt in OUTPUT_FORMATS:
            tk.Radiobutton(fr, text=fmt, variable=self.format_var, value=fmt,
                           bg=P.BG3, fg=P.OR, selectcolor=P.BG3, activebackground=P.BG3,
                           font=("Consolas",9,"bold")).pack(side="left", padx=8)
        qr = tk.Frame(body, bg=P.BG3); qr.pack(fill="x", padx=10, pady=(0,4))
        tk.Label(qr, text="QUALITY:", font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(side="left", padx=(0,8))
        ttk.Combobox(qr, textvariable=self.quality_var, values=list(QUALITY_PRESETS.keys()), state="readonly", font=P.MONO, width=22).pack(side="left")
        tk.Label(body, text="  CRF = Constant Rate Factor  |  lower = better quality  |  18 = near-lossless",
                 font=("Consolas",7), fg=P.TXT3, bg=P.BG3).pack(anchor="w", padx=10, pady=(0,2))
        af = tk.Frame(body, bg=P.BG3); af.pack(fill="x", padx=10, pady=(0,10))
        for lbl, var, opts, w in [("AUDIO:", self.audio_var, AUDIO_BITRATES, 10),
                                  ("FPS:",   self.fps_var,   FPS_OPTIONS,    13)]:
            tk.Label(af, text=lbl, font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(side="left", padx=(0,4))
            ttk.Combobox(af, textvariable=var, values=opts, state="readonly", font=P.MONO, width=w).pack(side="left", padx=(0,14))

    def _build_advanced_panel(self, p):
        body = self._card(p, "ADVANCED SETTINGS", P.TXT2)
        row = tk.Frame(body, bg=P.BG3); row.pack(fill="x", padx=10, pady=10)
        tk.Label(row, text="ENCODE SPEED:", font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(side="left", padx=(0,6))
        ttk.Combobox(row, textvariable=self.speed_var, values=SPEED_PRESETS, state="readonly", font=P.MONO, width=12).pack(side="left", padx=(0,14))
        for txt, var, col in [("Auto ZIP", self.autozip_var, P.G1), ("Skip Done", self.skip_done_var, P.B1)]:
            tk.Checkbutton(row, text=txt, variable=var,
                           font=P.MONO_S, fg=col, bg=P.BG3, selectcolor=P.BG3,
                           activebackground=P.BG3, activeforeground=col).pack(side="left", padx=(0,12))

    def _build_placement_panel(self, p):
        body = self._card(p, "PLACEMENT  MODE", P.PK)
        tk.Label(body, text="  Smart Crop: scale-to-fill then center-crop  (CSS cover / Photoshop fit)",
                 font=("Consolas",7), fg=P.G1, bg=P.BG3).pack(anchor="w", padx=10, pady=(6,2))
        for label, key, is_ai in PLACEMENT_MODES:
            f = tk.Frame(body, bg=P.BG3, cursor="hand2", highlightthickness=1, highlightbackground=P.BDR)
            f.pack(fill="x", padx=10, pady=2)
            inner = tk.Frame(f, bg=P.BG3); inner.pack(fill="x", padx=4, pady=4)
            rb = tk.Radiobutton(inner, variable=self.placement_var, value=key,
                                bg=P.BG3, activebackground=P.BG3, selectcolor=P.PK, fg=P.PK,
                                command=lambda k=key, fr=f: self._sel_placement(k, fr))
            rb.pack(side="left")
            tk.Label(inner, text=f"  {label}", font=P.MONO, fg=P.TXT, bg=P.BG3).pack(side="left")
            if is_ai:
                is_ok = HAS_YOLO if key == "yolo_crop" else True
                ai_col = P.PU if is_ok else P.RD
                ai_txt = " [YOLOv8]" if key == "yolo_crop" else " [AI]"
                tk.Label(inner, text=ai_txt, font=("Consolas",7,"bold"), fg=ai_col, bg=P.BG3).pack(side="left")
            self._placement_frames[key] = f
            for widget in ([f, inner] + list(inner.winfo_children())):
                try:
                    widget.bind("<Button-1>", lambda e, k=key, fr=f: (self.placement_var.set(k), self._sel_placement(k, fr)))
                except: pass
        self._sel_placement("smart_crop", self._placement_frames["smart_crop"])

    def _sel_placement(self, key, frame):
        try:
            for k, f in self._placement_frames.items():
                f.configure(bg=P.BG3, highlightbackground=P.BDR)
                for w in f.winfo_children():
                    try: w.configure(bg=P.BG3)
                    except: pass
                    for ww in (w.winfo_children() if hasattr(w,"winfo_children") else []):
                        try: ww.configure(bg=P.BG3)
                        except: pass
            frame.configure(bg="#1a0e30", highlightbackground=P.PK)
            for w in frame.winfo_children():
                try: w.configure(bg="#1a0e30")
                except: pass
                for ww in (w.winfo_children() if hasattr(w,"winfo_children") else []):
                    try: ww.configure(bg="#1a0e30")
                    except: pass
        except: pass

    def _build_shopify_panel(self, p):
        body = self._card(p, "SHOPIFY  SIZE  GUARD", P.OR)
        row = tk.Frame(body, bg=P.BG3); row.pack(fill="x", padx=10, pady=(10,4))
        tk.Checkbutton(row, text="Enable size limit", variable=self.shopify_var,
                       font=("Consolas",9,"bold"), fg=P.OR, bg=P.BG3, selectcolor=P.BG3,
                       activebackground=P.BG3, activeforeground=P.OR).pack(side="left")
        tk.Label(row, text="  Max:", font=P.MONO_S, fg=P.TXT3, bg=P.BG3).pack(side="left")
        ttk.Combobox(row, textvariable=self.max_mb_var, values=["5","7","10","15","20","50","100"],
                     state="readonly", font=P.MONO, width=5).pack(side="left")
        tk.Label(row, text=" MB", font=P.MONO_S, fg=P.TXT, bg=P.BG3).pack(side="left")
        tk.Label(body, text="  CRF primary  ->  true 2-pass if over limit  ->  faststart  ->  yuv420p",
                 font=("Consolas",7), fg=P.TXT3, bg=P.BG3).pack(anchor="w", padx=10, pady=(0,10))

    def _build_output_panel(self, p):
        body = self._card(p, "OUTPUT FOLDER  (subfolder structure mirrored)", P.G1)
        row = tk.Frame(body, bg=P.BG3); row.pack(fill="x", padx=10, pady=10)
        tk.Entry(row, textvariable=self.output_dir, font=P.MONO, fg=P.TXT, bg=P.BG0, insertbackground=P.G1,
                 relief="flat", highlightthickness=1, highlightbackground=P.BDR).pack(side="left", fill="x", expand=True, ipady=5)
        NeonBtn(row, " Browse ", self._browse_output, P.G1).pack(side="right", padx=(4,0))
        NeonBtn(row, " Open ",   self._open_output,   P.B1).pack(side="right", padx=(4,0))

    def _build_right_panel(self, p):
        tk.Label(p, text="  LIVE PREVIEW", font=P.MONO_B, fg=P.G1, bg=P.BG0).pack(anchor="w", pady=(8,3))
        self.preview_cv = tk.Canvas(p, bg=P.BG0, height=180, highlightthickness=1, highlightbackground=P.BDR)
        self.preview_cv.pack(fill="x", padx=8, pady=(0,2))
        self._preview_placeholder()
        self.preview_name = tk.Label(p, text="", font=("Consolas",7), fg=P.TXT3, bg=P.BG0)
        self.preview_name.pack(anchor="w", padx=10, pady=(0,4))
        tk.Frame(p, bg=P.BDR, height=1).pack(fill="x", padx=8, pady=(2,4))
        self.det_panel = DetectionPanel(p); self.det_panel.pack(fill="x")
        tk.Frame(p, bg=P.BDR, height=1).pack(fill="x", padx=8, pady=(2,4))
        tk.Label(p, text="  STATISTICS", font=P.MONO_B, fg=P.G1, bg=P.BG0).pack(anchor="w")
        sg = tk.Frame(p, bg=P.BG0); sg.pack(fill="x", padx=8, pady=(4,6))
        self.stat_lbl = {}
        for i,(key,lbl,col) in enumerate([("total","TOTAL",P.B1),("done","DONE",P.G1),("failed","FAILED",P.RD)]):
            f = tk.Frame(sg, bg=P.BG3, highlightthickness=1, highlightbackground=col)
            f.grid(row=0, column=i, padx=3, sticky="nsew")
            sg.columnconfigure(i, weight=1)
            tk.Label(f, text=lbl, font=P.MONO_S, fg=col, bg=P.BG3).pack(pady=(5,0))
            sl = tk.Label(f, text="0", font=("Consolas",22,"bold"), fg=col, bg=P.BG3); sl.pack(pady=(0,5))
            self.stat_lbl[key] = sl
        tk.Label(p, text="  PROGRESS", font=P.MONO_B, fg=P.G1, bg=P.BG0).pack(anchor="w", pady=(4,2))
        self.progress_bar = ttk.Progressbar(p, mode="determinate", maximum=100, style="Green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", padx=8, pady=(0,2))
        self.progress_txt = tk.Label(p, text="0 / 0  (0.0%)", font=P.MONO_S, fg=P.TXT3, bg=P.BG0)
        self.progress_txt.pack(anchor="w", padx=10)
        tk.Frame(p, bg=P.BDR, height=1).pack(fill="x", padx=8, pady=(6,2))
        lhdr = tk.Frame(p, bg=P.BG0); lhdr.pack(fill="x", padx=8, pady=(2,2))
        tk.Label(lhdr, text="SESSION LOG", font=P.MONO_B, fg=P.G1, bg=P.BG0).pack(side="left")
        NeonBtn(lhdr, "Clear", self._clear_log, P.TXT2).pack(side="right")
        tk.Label(lhdr, text=str(LOG_FILE.name), font=("Consolas",6), fg=P.TXT3, bg=P.BG0).pack(side="right", padx=8)
        self.log_box = tk.Text(p, bg=P.BG0, fg=P.TXT, font=("Consolas",7), height=12, relief="flat", wrap="word",
                               state="disabled", borderwidth=0, selectbackground=P.PU2)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(0,4))
        for tag, col in [("ok",P.G1),("error",P.RD),("warn",P.OR),("yolo",P.PU),("info",P.B1),("head",P.YL)]:
            self.log_box.tag_config(tag, foreground=col)
        tk.Frame(p, bg=P.BDR, height=1).pack(fill="x", padx=8, pady=(2,4))
        tk.Label(p, text="  AI LIBRARY STATUS", font=P.MONO_B, fg=P.G1, bg=P.BG0).pack(anchor="w")
        libs = [("YOLOv8", HAS_YOLO, P.PU, "pip install ultralytics"), ("OpenCV", HAS_CV2, P.B1, "pip install opencv-python"),
                ("moviepy", HAS_MOVIEPY, P.G1, "pip install moviepy"), ("Pillow", HAS_PIL, P.YL, "pip install Pillow"),
                ("ttk", HAS_TB, P.PK, "pip install ttkbootstrap"), ("ffmpeg", HAS_FFMPEG, P.CY, "Install from ffmpeg.org")]
        for name, ok, col, hint in libs:
            bg2 = "#041408" if ok else "#140408"
            f = tk.Frame(p, bg=bg2, highlightthickness=1, highlightbackground=col if ok else P.RD)
            f.pack(fill="x", padx=8, pady=1)
            suffix = "  READY" if ok else f"  -> {hint}"
            tk.Label(f, text=f"  {'OK' if ok else '!!'} {name}{suffix}", font=("Consolas",7), fg=col if ok else P.RD, bg=bg2, anchor="w").pack(fill="x", pady=2)

    def _build_bottombar(self):
        bar = tk.Frame(self.root, bg=P.BG0, height=56); bar.pack(fill="x", side="bottom"); bar.pack_propagate(False)
        self.start_btn = tk.Button(bar, text="   START PROCESSING   ", font=("Consolas",15,"bold"),
                                   fg=P.BG0, bg=P.G1, activebackground=P.G2, activeforeground=P.BG0,
                                   relief="flat", cursor="hand2", command=self._start)
        self.start_btn.pack(side="left", fill="both", expand=True, padx=(10,4), pady=8)
        tk.Button(bar, text="  STOP  ", font=("Consolas",14,"bold"), fg="#fff", bg=P.RD, activebackground=P.RD2,
                  relief="flat", cursor="hand2", command=self._stop, width=10).pack(side="right", fill="y", padx=(0,10), pady=8)
        self.status_lbl = tk.Label(bar, text="Ready.", font=P.MONO, fg=P.TXT3, bg=P.BG0); self.status_lbl.pack(side="right", padx=16)

    def _pulse_logo(self):
        try:
            self.logo_lbl.config(fg=self._LOGO_FRAMES[self._logo_idx])
            self._logo_idx = (self._logo_idx + 1) % len(self._LOGO_FRAMES)
            self.root.after(700, self._pulse_logo)
        except: pass

    def _preview_placeholder(self):
        try:
            self.preview_cv.delete("all")
            self.preview_cv.create_text(200, 90, text="No Preview", fill=P.TXT3, font=P.MONO_S)
        except: pass

    def _on_preset_change(self, _):
        try:
            w, h = PRESETS.get(self.preset_var.get(), (None, None))
            if w: self.width_var.set(str(w)); self.height_var.set(str(h))
        except: pass

    def _warn_ffmpeg(self):
        messagebox.showwarning("ffmpeg Not Found",
            "ffmpeg was not found in this app bundle.\n\n"
            "Please install ffmpeg:\n"
            "  brew install ffmpeg\n\n"
            "Or download from: https://ffmpeg.org/download.html\n"
            "Then add it to your PATH.")

    def _add_files(self):
        try:
            files = filedialog.askopenfilenames(title="Select Videos",
                filetypes=[("Video files", " ".join(f"*{e}" for e in VIDEO_EXTS)), ("All files","*.*")])
            if files:
                seen = {i["abs"] for i in self.input_files}
                for f in files:
                    if f not in seen: self.input_files.append({"abs": f, "rel": os.path.basename(f)})
                self._refresh_list()
        except Exception as exc: log.error("_add_files: %s", exc)

    def _add_folder(self):
        try:
            folder = filedialog.askdirectory(title="Select Root Folder (all subfolders scanned)")
            if not folder: return
            found = scan_folder_for_videos(folder)
            if not found:
                messagebox.showinfo("No Videos", "No video files found.")
                return
            seen = {i["abs"] for i in self.input_files}
            for item in found:
                if item["abs"] not in seen: self.input_files.append(item)
            self._refresh_list()
            subdirs = {str(Path(i["rel"]).parent) for i in self.input_files if str(Path(i["rel"]).parent) != "."}
            if subdirs: self.struct_lbl.config(text=f"  {len(subdirs)} subfolder(s) found - structure mirrored")
            else: self.struct_lbl.config(text="  Flat folder - no subfolders")
        except Exception as exc: log.error("_add_folder: %s", exc)

    def _clear_files(self):
        self.input_files.clear(); self.struct_lbl.config(text=""); self._refresh_list()

    def _refresh_list(self):
        try:
            self.file_lb.delete(0, "end")
            for item in self.input_files: self.file_lb.insert("end", f"  {item['rel']}")
            n = len(self.input_files)
            self.file_count_lbl.config(text=f"  {n} video{'s' if n!=1 else ''} queued", fg=P.G1 if n>0 else P.TXT3)
        except: pass

    def _browse_output(self):
        try:
            d = filedialog.askdirectory(title="Choose Output Folder")
            if d: self.output_dir.set(d)
        except: pass

    def _open_output(self):
        try:
            d = self.output_dir.get(); Path(d).mkdir(parents=True, exist_ok=True)
            if sys.platform == "darwin": subprocess.Popen(["open", d])
            elif sys.platform == "win32": os.startfile(d)
            else: subprocess.Popen(["xdg-open", d])
        except Exception as exc: log.error("_open_output: %s", exc)

    def _start(self):
        if not HAS_FFMPEG: self._warn_ffmpeg(); return
        if not self.input_files: messagebox.showwarning("No Files", "Add video files or a folder first."); return
        try:
            w = int(self.width_var.get()); h = int(self.height_var.get())
            if w <= 0 or h <= 0: raise ValueError
        except:
            messagebox.showerror("Size Error", "Enter valid Width and Height."); return
        if self.shopify_var.get():
            try: max_bytes = int(float(self.max_mb_var.get()) * 1024 * 1024)
            except: max_bytes = 7 * 1024 * 1024
        else: max_bytes = 999 * 1024 * 1024
        cfg = {"width": str(w), "height": str(h), "placement": self.placement_var.get(), "format": self.format_var.get(),
               "crf": QUALITY_PRESETS.get(self.quality_var.get(), "22"), "speed": self.speed_var.get(),
               "fps": self.fps_var.get(), "audio_bitrate": self.audio_var.get(), "output_dir": self.output_dir.get(),
               "auto_zip": self.autozip_var.get(), "skip_done": self.skip_done_var.get(), "max_bytes": max_bytes}
        self._clear_log(); self.det_panel.clear(); self.start_btn.config(state="disabled")
        self.status_lbl.config(text="Processing...", fg=P.G1)
        self.processor = VideoProcessor(log_cb=self._on_log, progress_cb=self._on_progress,
                                        preview_cb=self._on_preview, detection_cb=self._on_detection)
        self._proc_thread = threading.Thread(target=self._run_batch, args=(list(self.input_files), cfg), daemon=True)
        self._proc_thread.start()

    def _run_batch(self, files, cfg):
        try: self.processor.process_batch(files, cfg)
        except Exception as exc:
            log.error("_run_batch unhandled: %s", traceback.format_exc())
            self._on_log(f"CRITICAL ERROR: {exc}", "error")
        finally: self.root.after(0, self._on_done)

    def _on_done(self):
        try: self.start_btn.config(state="normal"); self.status_lbl.config(text="Done!", fg=P.G1)
        except: pass

    def _stop(self):
        try:
            if self.processor: self.processor.stop()
            self.status_lbl.config(text="Stopping...", fg=P.RD)
        except: pass

    def _on_close(self):
        try:
            if self.processor: self.processor.stop()
        except: pass
        try: self.root.destroy()
        except: pass

    def _on_log(self, msg, level="info"):
        def _do():
            try:
                self.log_box.config(state="normal")
                tag = level if level in ("ok","error","warn","yolo","info","head") else "info"
                self.log_box.insert("end", msg + "\n", tag)
                self.log_box.see("end")
                self.log_box.config(state="disabled")
            except: pass
        try: self.root.after(0, _do)
        except: pass

    def _clear_log(self):
        try: self.log_box.config(state="normal"); self.log_box.delete("1.0","end"); self.log_box.config(state="disabled")
        except: pass

    def _on_progress(self, current, total, done, failed):
        def _do():
            try:
                pct = (current / total * 100) if total else 0
                self.progress_bar["value"] = pct
                self.progress_txt.config(text=f"  {current} / {total}  ({pct:.1f}%)")
                self.stat_lbl["total"].config(text=str(total))
                self.stat_lbl["done"].config(text=str(done))
                self.stat_lbl["failed"].config(text=str(failed))
            except: pass
        try: self.root.after(0, _do)
        except: pass

    def _on_preview(self, video_path):
        def _do():
            if not HAS_CV2 or not HAS_PIL: return
            try:
                cap = cv2.VideoCapture(str(video_path))
                if not cap.isOpened(): return
                ret, frame = cap.read(); cap.release()
                if not ret: return
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(rgb)
                cw = max(self.preview_cv.winfo_width(), 390)
                img.thumbnail((cw, 180))
                imgtk = ImageTk.PhotoImage(img)
                self.preview_cv.delete("all")
                self.preview_cv.config(height=max(img.height, 40))
                self.preview_cv.create_image(cw//2, img.height//2, anchor="center", image=imgtk)
                self.preview_cv._img = imgtk
                sz = Path(video_path).stat().st_size / (1024*1024)
                limit = float(self.max_mb_var.get())
                col = P.G1 if sz <= limit else P.RD
                self.preview_name.config(text=f"  {Path(video_path).name}  --  {sz:.2f} MB", fg=col)
            except Exception as exc: log.debug("_on_preview error: %s", exc)
        try: self.root.after(0, _do)
        except: pass

    def _on_detection(self, frame_bgr, detections):
        def _do():
            try: self.det_panel.update_detection(frame_bgr, detections)
            except: pass
        try: self.root.after(0, _do)
        except: pass

def main():
    try:
        if HAS_TB:
            root = tb.Window(themename="darkly")
        else:
            root = tk.Tk()
        app = VideoResizerApp(root)
        root.mainloop()
    except Exception as exc:
        log.critical("Fatal error: %s", traceback.format_exc())
        try:
            import tkinter.messagebox as mb
            mb.showerror("Fatal Error", str(exc))
        except: pass
        sys.exit(1)

if __name__ == "__main__":
    main()
