import logging
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from time import perf_counter_ns, thread_time_ns
from tkinter import (  # noqa
    BOTH,
    END,
    LEFT,
    RIGHT,
    StringVar,
    Toplevel,
    Y,
    messagebox,
    ttk,
)
from typing import Dict, Iterable, List, Optional
from urllib.parse import unquote, urlparse

import imagehash
import pythoncom
from ttkthemes import ThemedTk
import win32com.client
import win32gui
from PIL import (
    Image,
    ImageFile,
    ImageFilter,
    ImageGrab,
    ImageOps,
    ImageStat,
)

if __name__ != "__main__":
    import modules.status

from modules.operations import build_search_string, search_with_everything, search_with_everything_file_list

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}

#region --------------------------- Profiling ---------------------------

@dataclass
class ProfileEntry:
    count: int = 0
    wall_ns: int = 0
    cpu_ns: int = 0


@dataclass
class Profiler:
    entries: Dict[str, ProfileEntry] = field(default_factory=dict)
    _lock: threading.Lock = field(default_factory=threading.Lock, repr=False)

    def record(self, label: str, wall_ns: int, cpu_ns: int) -> None:
        with self._lock:
            e = self.entries.get(label)
            if e is None:
                e = ProfileEntry()
                self.entries[label] = e
            e.count += 1
            e.wall_ns += wall_ns
            e.cpu_ns += cpu_ns

    def timer(self, label: str):
        profiler = self

        class _Timer:
            def __enter__(self_):
                self_.t0 = perf_counter_ns()
                self_.c0 = thread_time_ns()  # per-thread CPU, safe under threads
                return self_

            def __exit__(self_, exc_type, exc, tb):
                wall_ns = perf_counter_ns() - self_.t0
                cpu_ns = thread_time_ns() - self_.c0
                profiler.record(label, wall_ns, cpu_ns)

        return _Timer()

    def summary_lines(self, show_avg: bool = True) -> list[str]:
        # Sort by total wall desc
        items = sorted(self.entries.items(), key=lambda kv: kv[1].wall_ns, reverse=True)
        lines = []
        for k, v in items:
            wall_ms = v.wall_ns / 1e6
            cpu_ms = v.cpu_ns / 1e6
            if v.count and show_avg:
                aw = wall_ms / v.count
                ac = cpu_ms / v.count
                lines.append(f"{k}: {wall_ms:.1f} ms wall, {cpu_ms:.1f} ms cpu (n={v.count}, avg {aw:.2f}/{ac:.2f} ms)")
            else:
                lines.append(f"{k}: {wall_ms:.1f} ms wall, {cpu_ms:.1f} ms cpu (n={v.count})")
        return lines


def log_profile(prof: Profiler, header: str = "Fuzzy Image Search profile"):
    logging.info(header)
    lines = prof.summary_lines()
    for line in lines:
        logging.info("  " + line)

    # Optional: add a concurrency factor for key sections
    def _ns(label):
        e = prof.entries.get(label)
        return (e.wall_ns, e.count) if e else (0, 0)

    stage_wall_ns, _ = _ns("prefilter.total")
    if stage_wall_ns:
        qm_wall_ns, qm_n = _ns("prefilter.quick_metrics")
        if qm_n:
            cf = qm_wall_ns / stage_wall_ns  # “sum of per-file wall” / “stage elapsed”
            avg_ms = (qm_wall_ns / qm_n) / 1e6
            logging.info(
                f"  quick_metrics throughput: {qm_n} files in {stage_wall_ns/1e6:.1f} ms "
                f"({(qm_n / (stage_wall_ns/1e9)):.1f} files/s), "
                f"avg {avg_ms:.2f} ms per file, concurrency factor {cf:.1f}x"
            )


#region --------------------------- Clipboard ---------------------------


def _load_first_image_from_paths(paths: Iterable[str]) -> Optional[Image.Image]:
    for p in paths:
        if Path(p).suffix.lower() in IMAGE_EXTS:
            try:
                return Image.open(p)
            except Exception:
                continue
    return None


def get_clipboard_image() -> Image.Image:
    ref = ImageGrab.grabclipboard()
    if isinstance(ref, Image.Image):
        return ref
    if isinstance(ref, list):
        img = _load_first_image_from_paths(ref)
        if img is not None:
            return img
    raise RuntimeError("No image on clipboard. Copy an image (bitmap or image file) and try again.")


#region --------------------------- Focused Explorer Folder ---------------------------


def _file_url_to_path(url: str) -> Path:
    parsed = urlparse(url)
    path = unquote(parsed.path)
    if parsed.netloc:
        path = f"//{parsed.netloc}{path}"
    if len(path) >= 3 and path[0] == "/" and path[2] == ":":
        path = path[1:]
    return Path(path)


def get_focused_explorer_folder() -> Path:
    pythoncom.CoInitialize()  # balanced by interpreter exit; short-lived call
    hwnd_fore = win32gui.GetForegroundWindow()
    shell = win32com.client.Dispatch("Shell.Application")
    # Prefer exact foreground match
    for w in shell.Windows():
        try:
            if int(w.HWND) == int(hwnd_fore) and str(w.LocationURL).lower().startswith("file:"):
                return _file_url_to_path(str(w.LocationURL))
        except Exception:
            continue
    # Fallback: any visible explorer window
    for w in shell.Windows():
        try:
            if str(w.FullName).lower().endswith("explorer.exe") and str(w.LocationURL).lower().startswith("file:"):
                return _file_url_to_path(str(w.LocationURL))
        except Exception:
            continue
    raise RuntimeError("Could not determine the focused File Explorer folder.")


#region --- Quick prefilter ---

ImageFile.LOAD_TRUNCATED_IMAGES = True
_ORIENT_TAG = 274  # EXIF orientation
_ROTATES_90 = {5, 6, 7, 8}

# knobs
QUICK_THUMB = 32  # default 32
EDGE_ON = False  # disable for speed-first
EDGE_THUMB = 12


def _get_orientation(im: Image.Image) -> int:
    try:
        ex = im.getexif()
        if ex:
            return int(ex.get(_ORIENT_TAG, 1))
    except Exception:
        pass
    return 1


def _oriented_wh(wh: tuple[int, int], orientation: int) -> tuple[int, int]:
    w, h = wh
    return (h, w) if orientation in _ROTATES_90 else (w, h)


# Apply orientation to a small image only (avoid full-res transpose)
def _apply_orientation_small(im_small: Image.Image, orientation: int) -> Image.Image:
    # Map EXIF orientations to Pillow ops
    if orientation == 1:
        return im_small
    if orientation == 2:
        return ImageOps.mirror(im_small)
    if orientation == 3:
        return im_small.rotate(180, expand=True)
    if orientation == 4:
        return ImageOps.flip(im_small)
    if orientation == 5:
        return ImageOps.mirror(im_small.rotate(270, expand=True))
    if orientation == 6:
        return im_small.rotate(270, expand=True)
    if orientation == 7:
        return ImageOps.mirror(im_small.rotate(90, expand=True))
    if orientation == 8:
        return im_small.rotate(90, expand=True)
    return im_small


def _aspect_ratio(size):
    w, h = size
    return (w / h) if h else 0.0


@dataclass
class QuickMetrics:
    size: tuple[int, int]
    aspect: float
    mean_rgb: tuple[float, float, float]
    lum_mean: float
    lum_var: float
    edge_mean: float


def _compute_quick_metrics(img: Image.Image, thumb=QUICK_THUMB) -> QuickMetrics:
    # DO NOT exif_transpose(img) here — that can decode full-res.
    # Read orientation tag cheaply, then apply to the small thumb.
    orientation = _get_orientation(img)

    # Make a tiny thumbnail directly from the lazy-loaded image.
    im2 = img.copy()  # this may still trigger decode; keep it tiny
    im2.thumbnail((thumb, thumb), resample=Image.BILINEAR)  # forces decode at small size

    # Now apply orientation to the already-small image
    im2 = _apply_orientation_small(im2, orientation)

    if im2.mode != "RGB":
        im2 = im2.convert("RGB")

    # Use oriented header dims (swap for 90° rotations); no full-res decode.
    hdr_size = getattr(img, "size", im2.size)
    size = _oriented_wh(hdr_size, orientation)
    aspect = _aspect_ratio(size)

    # Stats on tiny image
    stat = ImageStat.Stat(im2)
    mean_rgb = tuple(float(x) for x in stat.mean[:3])

    lum = im2.convert("L")
    st_l = ImageStat.Stat(lum)
    lum_mean = float(st_l.mean[0])
    lum_var = float(st_l.var[0])

    edges = lum.filter(ImageFilter.FIND_EDGES)
    edge_mean = float(ImageStat.Stat(edges).mean[0])

    return QuickMetrics(
        size=size,
        aspect=aspect,
        mean_rgb=mean_rgb,
        lum_mean=lum_mean,
        lum_var=lum_var,
        edge_mean=edge_mean,
    )


def _rgb_distance(a: tuple[float, float, float], b: tuple[float, float, float]) -> float:
    # Simple Euclidean in sRGB mean space (fast, not perceptual)
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2) ** 0.5


@dataclass
class PrefilterParams:
    min_bytes: int = 8 * 1024  # ignore tiny icons/empties
    max_bytes: int = 50 * 1024 * 1024  # cap huge files
    min_w: int = 64
    min_h: int = 64
    max_w: int = 20000
    max_h: int = 20000
    aspect_tol: float = 0.20  # ±20% aspect window vs ref
    color_max_delta: float = 50.0  # mean RGB distance threshold (0–255 scale)
    lum_var_tol: float = 0.7  # keep if ratio in [tol..1/tol]
    edge_mean_tol: float = 0.65  # same idea as above
    # Optional pre-aHash stage
    ahash_size: int = 8
    ahash_max_dist: int = 20  # reject if aHash Hamming > 16

    @classmethod
    def testing(cls):
        p = cls()
        p.min_bytes = 1
        p.max_bytes = 10**9
        p.color_max_delta = 1e6
        p.lum_var_tol = 0.1
        p.edge_mean_tol = 0.1
        p.ahash_max_dist = 64
        return p


def _within_ratio(a: float, b: float, tol: float) -> bool:
    if a <= 0 or b <= 0:
        return False
    r = a / b
    return (1.0 - tol) <= r <= (1.0 + tol)


def _compute_quick_metrics_from_small(im_small: Image.Image, orig_wh: tuple[int, int]) -> QuickMetrics:
    # im_small is small, oriented, RGB
    stat = ImageStat.Stat(im_small)
    mean_rgb = tuple(float(x) for x in stat.mean[:3])

    lum = im_small.convert("L")
    st_l = ImageStat.Stat(lum)
    lum_mean = float(st_l.mean[0])
    lum_var = float(st_l.var[0])

    if EDGE_ON:
        # ultra-cheap edges: even smaller and 8-bit ops
        e = lum.copy()
        e.thumbnail((EDGE_THUMB, EDGE_THUMB), resample=Image.BILINEAR)
        e = e.filter(ImageFilter.FIND_EDGES)
        edge_mean = float(ImageStat.Stat(e).mean[0])
    else:
        edge_mean = 0.0  # neutral; adjust tolerance logic to ignore if 0

    return QuickMetrics(
        size=orig_wh,
        aspect=_aspect_ratio(orig_wh),
        mean_rgb=mean_rgb,
        lum_mean=lum_mean,
        lum_var=lum_var,
        edge_mean=edge_mean,
    )


def _passes_quick_filters_path(
    path: Path,
    ref_qm: QuickMetrics,
    p: PrefilterParams,
    profiler: Optional[Profiler] = None,
) -> bool:
    # OS stat
    try:
        if profiler:
            t0 = perf_counter_ns()
            c0 = thread_time_ns()
        st = path.stat()
        if profiler:
            profiler.record("prefilter.stat", perf_counter_ns() - t0, thread_time_ns() - c0)
    except OSError:
        return False
    if st.st_size < p.min_bytes or st.st_size > p.max_bytes:
        return False

    # One open only
    try:
        t0 = perf_counter_ns()
        c0 = thread_time_ns()
        with Image.open(path) as im:
            if profiler:
                profiler.record(
                    "prefilter.open_header",
                    perf_counter_ns() - t0,
                    thread_time_ns() - c0,
                )

            # Cheap header checks
            w, h = im.size
            orientation = _get_orientation(im)
            w_o, h_o = _oriented_wh((w, h), orientation)
            if not (p.min_w <= w_o <= p.max_w and p.min_h <= h_o <= p.max_h):
                return False
            if not _within_ratio(_aspect_ratio((w_o, h_o)), ref_qm.aspect, p.aspect_tol):
                return False

            # Build tiny, oriented RGB thumb ONCE
            if profiler:
                with profiler.timer("prefilter.quick_metrics"):
                    # create thumb and derive metrics
                    im_small = im.copy()
                    im_small.thumbnail((QUICK_THUMB, QUICK_THUMB), resample=Image.BILINEAR)
                    im_small = _apply_orientation_small(im_small, orientation)
                    if im_small.mode != "RGB":
                        t1 = perf_counter_ns()
                        cc1 = thread_time_ns()
                        im_small = im_small.convert("RGB")
                        profiler.record(
                            "prefilter.convert_rgb",
                            perf_counter_ns() - t1,
                            thread_time_ns() - cc1,
                        )
                    qm = _compute_quick_metrics_from_small(im_small, (w, h))
            else:
                im_small = im.copy()
                im_small.thumbnail((QUICK_THUMB, QUICK_THUMB), resample=Image.BILINEAR)
                im_small = _apply_orientation_small(im_small, orientation)
                if im_small.mode != "RGB":
                    im_small = im_small.convert("RGB")
                qm = _compute_quick_metrics_from_small(im_small, (w, h))

            # aHash gate from the same tiny RGB image — no second open
            ah1 = None
            if p.ahash_max_dist >= 0:
                if profiler:
                    with profiler.timer("prefilter.ahash_compute"):
                        ah1 = imagehash.average_hash(im_small, hash_size=p.ahash_size)
                else:
                    ah1 = imagehash.average_hash(im_small, hash_size=p.ahash_size)
    except Exception:
        return False

    # Compare metrics
    if _rgb_distance(qm.mean_rgb, ref_qm.mean_rgb) > p.color_max_delta:
        return False
    r_lv = max(qm.lum_var, 1e-6) / max(ref_qm.lum_var, 1e-6)
    if not (p.lum_var_tol <= r_lv <= (1.0 / p.lum_var_tol)):
        return False

    # Edge density similarity (only if enabled)
    if EDGE_ON:
        r_em = max(qm.edge_mean, 1e-6) / max(ref_qm.edge_mean, 1e-6)
        if not (p.edge_mean_tol <= r_em <= (1.0 / p.edge_mean_tol)):
            return False

    if p.ahash_max_dist >= 0 and (ref_qm._ahash - ah1) > p.ahash_max_dist:
        return False

    return True


#region --------------------------- Hashing ---------------------------


@dataclass
class HashSignature:
    phash: imagehash.ImageHash
    dhash: imagehash.ImageHash
    bits_phash: int
    bits_dhash: int

    def distance_to(self, other: "HashSignature", w_phash: float, w_dhash: float):
        d_p = self.phash - other.phash
        d_d = self.dhash - other.dhash
        n_p = d_p / float(self.bits_phash)
        n_d = d_d / float(self.bits_dhash)
        weighted = w_phash * n_p + w_dhash * n_d
        denom = (w_phash + w_dhash) if (w_phash + w_dhash) > 0 else 1.0
        norm = weighted / denom
        return norm, {"phash": d_p, "dhash": d_d}


def compute_signature(img: Image.Image, hash_size: int = 16, profiler: Optional[Profiler] = None) -> HashSignature:
    if profiler:
        with profiler.timer("hash.transpose"):
            img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            with profiler.timer("hash.convert_rgb"):
                img = img.convert("RGB")
        with profiler.timer("hash.phash"):
            p = imagehash.phash(img, hash_size=hash_size)
        with profiler.timer("hash.dhash"):
            d = imagehash.dhash(img, hash_size=hash_size)
    else:
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        p = imagehash.phash(img, hash_size=hash_size)
        d = imagehash.dhash(img, hash_size=hash_size)
    bits = hash_size * hash_size
    return HashSignature(p, d, bits, bits)


def iter_image_files(folder: Path, recursive: bool = False):
    it = folder.rglob("*") if recursive else folder.glob("*")
    for p in it:
        if p.is_file() and p.suffix.lower() in IMAGE_EXTS:
            yield p


@dataclass
class MatchResult:
    path: Path
    similarity: float
    norm_distance: float
    breakdown: Dict[str, int]
    size_bytes: int
    dimensions: tuple[int, int]


def score_one(
    path: Path,
    ref_sig: HashSignature,
    hash_size: int,
    w_phash: float,
    w_dhash: float,
    profiler: Optional[Profiler] = None,
) -> Optional[MatchResult]:
    try:
        t0 = perf_counter_ns()
        c0 = thread_time_ns()
        img = Image.open(path)
        if profiler:
            profiler.record("hash.open_image", perf_counter_ns() - t0, thread_time_ns() - c0)
        try:
            img.seek(0)
        except Exception:
            pass
        w, h = img.size
        st_size = path.stat().st_size
    except Exception:
        return None
    sig = compute_signature(img, hash_size=hash_size, profiler=profiler)
    norm_dist, breakdown = ref_sig.distance_to(sig, w_phash, w_dhash)
    sim = 1.0 - norm_dist
    return MatchResult(path, sim, norm_dist, breakdown, st_size, (w, h))


def search(
    folder: Path,
    reference_img: Image.Image,
    recursive: bool = True,
    hash_size: int = 16,
    w_phash: float = 4.0,
    w_dhash: float = 2.0,
    max_files: Optional[int] = None,
    pre: Optional[PrefilterParams] = None,
    profiler: Optional[Profiler] = None,
):
    prof = profiler or Profiler()
    pre = pre or PrefilterParams()

    with prof.timer("ref.quick_metrics"):
        ref_qm = _compute_quick_metrics(reference_img)

    with prof.timer("ref.ahash"):
        ref_ahash = imagehash.average_hash(
            ImageOps.exif_transpose(
                reference_img if reference_img.mode in ("RGB", "L") else reference_img.convert("RGB")
            ),
            hash_size=pre.ahash_size,
        )
    setattr(ref_qm, "_ahash", ref_ahash)

    # Collect candidates (fast scandir)
    def _iter_images(folder: Path):
        stack = [folder]
        while stack:
            cur = stack.pop()
            with os.scandir(cur) as it:
                for e in it:
                    try:
                        if e.is_dir(follow_symlinks=False) and recursive:
                            stack.append(Path(e.path))
                        elif e.is_file(follow_symlinks=False):
                            ext = os.path.splitext(e.name)[1].lower()
                            if ext in IMAGE_EXTS:
                                yield Path(e.path)
                    except OSError:
                        continue

    with prof.timer("scan.walk"):
        items = list(_iter_images(folder))
    if max_files:
        items = items[:max_files]

    # Quick multi-threaded prefilter
    filtered: list[Path] = []
    with prof.timer("prefilter.total"):
        max_workers = min(32, (os.cpu_count() or 4) * 2)  # IO-bound
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futs = {ex.submit(_passes_quick_filters_path, pth, ref_qm, pre, prof): pth for pth in items}
            for fut in as_completed(futs):
                try:
                    if fut.result():
                        filtered.append(futs[fut])
                except Exception:
                    # ignore bad files; profiler already counted timings
                    pass

    logging.info(f"Prefilter: {len(filtered)} candidates out of {len(items)} files.")

    # Hashing stage only on survivors
    with prof.timer("ref.hash_signature"):
        ref_sig = compute_signature(reference_img, hash_size=hash_size, profiler=prof)

    results: list[MatchResult] = []
    with prof.timer("hashing.total"):
        for p in filtered:
            res = score_one(p, ref_sig, hash_size, w_phash, w_dhash, profiler=prof)
            if res is not None:
                results.append(res)

    with prof.timer("sort.results"):
        results.sort(key=lambda r: (-r.similarity, str(r.path).lower()))

    return results, prof


#region --------------------------- Tk UI ---------------------------


def _human_similarity(p: float) -> str:
    return f"{round(p * 100, 1)}%"


def run_dialog(
    top_n: int = 5,
    threshold: float = 0.55,
    recursive: bool = True,
    hash_size: int = 16,
    w_phash: float = 4.0,
    w_dhash: float = 2.0,
):
    """Grabs clipboard image, detects focused folder, searches, and shows a results dialog."""
    try:
        modules.status.update_status_message("Fuzzy search: reading clipboard...")
        t0 = perf_counter_ns()
        c0 = thread_time_ns()
        ref_img = get_clipboard_image()
        folder = get_focused_explorer_folder()
        setup_wall = perf_counter_ns() - t0
        setup_cpu = thread_time_ns() - c0
    except Exception as e:
        logging.exception("Setup failed: %s", str(e))
        # messagebox.showerror(
        #     "Fuzzy Image Search",
        #     str(e),
        #     parent=modules.status.root if modules.status.root else None,
        # )
        modules.status.update_status_message(f"Fuzzy search: Error: {str(e)}...")

        return

    modules.status.update_status_message(f"Fuzzy search: scanning {folder}...")
    try:
        results, prof = search(
            folder,
            ref_img,
            recursive=recursive,
            hash_size=hash_size,
            w_phash=w_phash,
            w_dhash=w_dhash,
        )
    except Exception as e:
        logging.exception("Search failed: %s", str(e))
        # messagebox.showerror(
        #     "Fuzzy Image Search",
        #     str(e),
        #     parent=modules.status.root if modules.status.root else None,
        # )
        modules.status.update_status_message(f"Fuzzy search: Error: {str(e)}...")
        return

    kept = [r for r in results if r.similarity >= threshold]
    if not kept:
        kept = results[:top_n]
        msg = f"No matches >= similarity threshold {threshold:.2f}."
        modules.status.update_status_message(f"Fuzzy search: {msg}")
        search_with_everything("No matches found")
        return
    else:
        kept = kept[:top_n]
        msg = f"Top {len(kept)} match(es)."
        search_with_everything_file_list(kept)
        modules.status.update_status_message(f"Fuzzy search: {msg}")


    # Build UI (timed)
    prof.record("setup.clipboard_and_folder", setup_wall, setup_cpu)
    log_profile(prof)

    root = modules.status.root

    def _clear_children(parent):
        for child in list(parent.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass

    with prof.timer("ui.build_window"):
        win = root
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        try:
            win.lift()
            win.focus_force()
            win.update()
        except Exception:
            pass

        # Drop "topmost" after a short delay while keeping focus
        def _drop_topmost():
            try:
                win.attributes("-topmost", False)
            except Exception:
                pass

        win.after(400, _drop_topmost)


        # Use a stable container so repeated runs replace the previous UI
        parent = modules.status.mainframe or win
        if modules.status.fuzzy_results_frame is None or not modules.status.fuzzy_results_frame.winfo_exists():
            modules.status.fuzzy_results_frame = ttk.Frame(parent)
            modules.status.fuzzy_results_frame.pack(fill=BOTH, expand=True)
        container = modules.status.fuzzy_results_frame
        _clear_children(container)

        # Treeview
        cols = ("similarity", "phash", "dhash", "dimensions", "size", "path")
        tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        widths = (80, 60, 60, 90, 70, 520)
        for c, w in zip(cols, widths):
            tree.heading(c, text=c.capitalize())
            tree.column(c, width=w, anchor="w")
        tree.pack(fill=BOTH, expand=True, padx=8, pady=8)

        def sortby(col, descending):
            data = [(tree.set(child, col), child) for child in tree.get_children("")]
            data.sort(reverse=descending)
            for i, item in enumerate(data):
                tree.move(item[1], "", i)
            tree.heading(col, command=lambda: sortby(col, not descending))

        for c in cols:
            tree.heading(c, command=lambda _c=c: sortby(_c, False))

        for r in kept:
            w, h = r.dimensions
            size_kb = round(r.size_bytes / 1024, 1)
            tree.insert(
                "",
                END,
                values=(
                    _human_similarity(r.similarity),
                    r.breakdown["phash"],
                    r.breakdown["dhash"],
                    f"{w}x{h}",
                    f"{size_kb} KB",
                    str(r.path),
                ),
            )

        # Actions
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill="x", padx=8, pady=8)
        open_btn = ttk.Button(btn_frame, text="Open Selected", command=lambda: _open_selected(tree))
        open_btn.pack(side=LEFT)
        reveal_btn = ttk.Button(btn_frame, text="Reveal in Explorer", command=lambda: _reveal_selected(tree))
        reveal_btn.pack(side=LEFT, padx=6)
        copy_btn = ttk.Button(btn_frame, text="Copy Path", command=lambda: _copy_selected(tree))
        copy_btn.pack(side=LEFT)

        def on_double(event):
            _open_selected(tree)

        tree.bind("<Double-1>", on_double)


def _selected_path(tree) -> Optional[Path]:
    sel = tree.selection()
    if not sel:
        return None
    vals = tree.item(sel[0], "values")
    if not vals or len(vals) < 4:
        return None
    return Path(vals[5])


def _open_selected(tree):
    p = _selected_path(tree)
    if p:
        os.startfile(str(p))


def _reveal_selected(tree):
    p = _selected_path(tree)
    if p:
        # Open Explorer and select the file
        os.system(f'explorer /select,"{str(p)}"')


def _copy_selected(tree):
    try:
        import pyperclip

        p = _selected_path(tree)
        if p:
            pyperclip.copy(str(p))
    except Exception:
        # silently ignore if pyperclip not available
        pass


if __name__ == "__main__":
    import sys

    arg_len = len(sys.argv)

    if arg_len == 2:
        ref_path = "<clipboard>"
        test_path = Path(sys.argv[1])
    elif arg_len == 3:
        ref_path = Path(sys.argv[1])
        test_path = Path(sys.argv[2])
    else:
        print(
            """Usage:
    Reference image refers to the image you'd like to check for duplicates.
    Candidate image is the file to evaluate against the reference for likeness.
    If reference image path is omitted, the image is instead taken from the last clipboard entry.
    python modules/fuzzy_image_search.py ref.jpg candidate.jpg
    python modules/fuzzy_image_search.py candidate.jpg
"""
        )
        sys.exit(1)

    ref_img = get_clipboard_image() if ref_path == "<clipboard>" else Image.open(ref_path)
    qm = _compute_quick_metrics(ref_img)
    qm._ahash = imagehash.average_hash(ref_img, hash_size=PrefilterParams().ahash_size)

    def prefilter_debug(path: Path, ref_qm: QuickMetrics, p: PrefilterParams):
        try:
            st = path.stat()
            if st.st_size < p.min_bytes or st.st_size > p.max_bytes:
                return "bytes"
            with Image.open(path) as im:
                w, h = im.size
                ori = _get_orientation(im)
                w_o, h_o = _oriented_wh((w, h), ori)
                if not (p.min_w <= w_o <= p.max_w and p.min_h <= h_o <= p.max_h):
                    return "dims"
                if not _within_ratio(_aspect_ratio((w_o, h_o)), ref_qm.aspect, p.aspect_tol):
                    return "aspect"

                im_small = im.copy()
                im_small.thumbnail((QUICK_THUMB, QUICK_THUMB), Image.BILINEAR)
                im_small = _apply_orientation_small(im_small, ori)
                if im_small.mode != "RGB":
                    im_small = im_small.convert("RGB")

                qm = _compute_quick_metrics_from_small(im_small, (w_o, h_o))
                if _rgb_distance(qm.mean_rgb, ref_qm.mean_rgb) > p.color_max_delta:
                    return "color"

                r_lv = max(qm.lum_var, 1e-6) / max(ref_qm.lum_var, 1e-6)
                if not (p.lum_var_tol <= r_lv <= (1.0 / p.lum_var_tol)):
                    return "lum_var"

                if p.ahash_max_dist >= 0:
                    ah1 = imagehash.average_hash(im_small, hash_size=p.ahash_size)
                    if (ref_qm._ahash - ah1) > p.ahash_max_dist:
                        return "ahash"

                return "pass"
        except Exception as e:
            return f"error:{type(e).__name__}"

    # Run diagnostics
    p = PrefilterParams()

    if test_path.is_dir():
        print(f"Scanning directory: {test_path}")
        total = 0
        summary = {}
        for img_path in iter_image_files(test_path, recursive=True):
            reason = prefilter_debug(img_path, qm, p)
            print(f"{reason}\t:\t{img_path.name}")
            total += 1
            summary[reason] = summary.get(reason, 0) + 1
        print("\nSummary:")
        for k, v in sorted(summary.items(), key=lambda kv: kv[1], reverse=True):
            print(f"  {k:10s} {v}")
        print(f"  total      {total}")
    else:
        reason = prefilter_debug(test_path, qm, p)
        print(f"{test_path.name}: {reason}")
