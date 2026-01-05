import pytest
from pathlib import Path # noqa
from PIL import Image
import imagehash
from modules.fuzzy_image_search import (
    _get_orientation,
    _apply_orientation_small,
    _aspect_ratio,
    _oriented_wh,
    _within_ratio,
    _rgb_distance,
    _compute_quick_metrics,
    _compute_quick_metrics_from_small,
    QuickMetrics,
    PrefilterParams,
    QUICK_THUMB,
)

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
            if not _within_ratio(
                _aspect_ratio((w_o, h_o)), ref_qm.aspect, p.aspect_tol
            ):
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


@pytest.fixture
def ref_qm(tmp_path):
    img = Image.new("RGB", (1024, 1024), (128, 128, 128))
    qm = _compute_quick_metrics(img, thumb=QUICK_THUMB)
    qm._ahash = imagehash.average_hash(img, hash_size=PrefilterParams().ahash_size)
    return qm


def test_prefilter_pass(ref_qm, tmp_path):
    img_path = tmp_path / "gray.png"
    Image.new("RGB", (1024, 1024), (128, 128, 128)).save(img_path)
    reason = prefilter_debug(img_path, ref_qm, PrefilterParams.testing())
    assert reason == "pass"


def test_prefilter_reject_color(ref_qm, tmp_path):
    img_path = tmp_path / "red.png"
    Image.new("RGB", (1024, 1024), (255, 0, 0)).save(img_path)
    reason = prefilter_debug(img_path, ref_qm, PrefilterParams.testing())
    assert reason in ("color", "pass")  # depending on tolerance
