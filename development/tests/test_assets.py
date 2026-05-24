from pathlib import Path

from PIL import Image

_ASSETS = Path(__file__).parents[1] / "app" / "assets"
_REQUIRED_ICO_SIZES = {(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)}


def _ico_sizes(image: Image.Image) -> set[tuple[int, int]]:
    ico = getattr(image, "ico", None)
    if ico is not None and hasattr(ico, "sizes"):
        return set(ico.sizes())

    sizes: set[tuple[int, int]] = set()
    for frame_index in range(getattr(image, "n_frames", 1)):
        image.seek(frame_index)
        sizes.add(image.size)
    return sizes


def test_logo_png_is_256_rgba():
    logo = _ASSETS / "filings_atlas_logo.png"

    assert logo.exists()
    with Image.open(logo) as image:
        assert image.size == (256, 256)
        assert image.mode == "RGBA"


def test_ico_has_exact_standard_sizes():
    ico = _ASSETS / "filings_atlas.ico"

    assert ico.exists()
    with Image.open(ico) as image:
        sizes = _ico_sizes(image)
        assert sizes == _REQUIRED_ICO_SIZES

        ico_reader = getattr(image, "ico", None)
        if ico_reader is not None and hasattr(ico_reader, "getimage"):
            for size in _REQUIRED_ICO_SIZES:
                assert ico_reader.getimage(size).size == size
