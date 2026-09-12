"""Offline image worker for the Lab Photo Tools PowerPoint add-in.

The only public IPC endpoint is ``python worker.py --request absolute.json``.
Image processing never downloads models or sends an image over a network.
"""
from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from dataclasses import dataclass
import hashlib
import json
import math
import os
from pathlib import Path
import sys
import tempfile
import warnings

from PIL import Image, ImageChops, ImageOps, UnidentifiedImageError

MAX_PIXELS = 40_000_000
MAX_ITEMS = 50
MAX_REQUEST_BYTES = 1_048_576
MODEL_FILENAME = "u2netp.onnx"
# Published by the rembg U2netpSession; not a user-configurable URL/model.
MODEL_MD5 = "8e83ca70e441ab06c318d82300c84806"
Image.MAX_IMAGE_PIXELS = MAX_PIXELS


class WorkerError(Exception):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


@dataclass(frozen=True)
class Item:
    input: Path
    output: Path
    angle: float


def validate_angle(value: object) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise WorkerError("INVALID_ANGLE", "각도는 -180°~180° 사이의 숫자로 입력해 주세요.")
    try:
        angle = float(value)
    except (OverflowError, ValueError):
        angle = float("nan")
    if not math.isfinite(angle) or not -180 <= angle <= 180:
        raise WorkerError("INVALID_ANGLE", "각도는 -180°~180° 사이의 유한한 숫자여야 합니다.")
    return angle


def absolute_path(value: object, label: str) -> Path:
    if not isinstance(value, str) or not value.strip() or "\x00" in value:
        raise WorkerError("INVALID_PATH", f"{label} 경로가 올바르지 않습니다.")
    path = Path(value)
    # UNC paths can send photographs to a file server; this worker is local-only.
    if not path.is_absolute() or str(path).startswith(("\\\\", "//")):
        raise WorkerError("INVALID_PATH", f"{label}에는 이 PC의 절대 경로를 사용해 주세요.")
    resolved = path.resolve()
    if str(resolved).startswith(("\\\\", "//")):
        raise WorkerError("INVALID_PATH", f"{label}에는 이 PC의 절대 경로를 사용해 주세요.")
    return resolved


def parse_request(data: object) -> tuple[str, list[Item], Path]:
    if not isinstance(data, dict):
        raise WorkerError("INVALID_REQUEST", "요청 형식이 올바르지 않습니다.")
    operation = data.get("operation")
    if operation not in ("straighten", "remove_background", "both"):
        raise WorkerError("INVALID_OPERATION", "지원하지 않는 사진 작업입니다.")
    raw_items = data.get("items")
    if not isinstance(raw_items, list) or not 1 <= len(raw_items) <= MAX_ITEMS:
        raise WorkerError("INVALID_ITEMS", "한 번에 사진 1~50개를 선택해 주세요.")
    items = []
    for entry in raw_items:
        if not isinstance(entry, dict):
            raise WorkerError("INVALID_ITEMS", "사진 항목 형식이 올바르지 않습니다.")
        source = absolute_path(entry.get("input"), "입력 사진")
        destination = absolute_path(entry.get("output"), "결과 사진")
        if not source.is_file():
            raise WorkerError("INPUT_MISSING", "선택한 사진의 임시 파일을 찾을 수 없습니다. 다시 시도해 주세요.")
        if destination.suffix.lower() != ".png":
            raise WorkerError("INVALID_OUTPUT", "결과 파일은 PNG 형식이어야 합니다.")
        if destination.exists() or not destination.parent.is_dir():
            raise WorkerError("OUTPUT_UNAVAILABLE", "결과 경로에 파일이 이미 있거나 폴더가 없습니다.")
        items.append(Item(source, destination, validate_angle(entry.get("angle", 0))))
    sources = {item.input for item in items}
    destinations = [item.output for item in items]
    if len(set(destinations)) != len(items) or any(path in sources for path in destinations):
        raise WorkerError("INVALID_OUTPUT", "원본과 결과 파일은 서로 다른 고유한 경로를 사용해야 합니다.")
    model_dir = absolute_path(data.get("model_dir") or os.getenv("U2NET_HOME") or str(Path(__file__).resolve().parent / "models"), "모델 폴더")
    return operation, items, model_dir


def load_image(path: Path) -> Image.Image:
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(path) as image:
                if image.format not in {"PNG", "JPEG", "BMP", "TIFF", "WEBP"}:
                    raise WorkerError("INVALID_IMAGE", "지원하는 사진 형식은 PNG, JPEG, BMP, TIFF, WEBP입니다.")
                if image.width < 1 or image.height < 1 or image.width * image.height > MAX_PIXELS:
                    raise WorkerError("IMAGE_TOO_LARGE", "사진 한 장은 4천만 화소 이하여야 합니다.")
                if getattr(image, "n_frames", 1) != 1:
                    raise WorkerError("MULTIFRAME_IMAGE", "여러 프레임이 있는 사진은 지원하지 않습니다. 한 장의 PNG로 변환해 주세요.")
                return ImageOps.exif_transpose(image).convert("RGBA")
    except (Image.DecompressionBombError, Image.DecompressionBombWarning):
        raise WorkerError("IMAGE_TOO_LARGE", "사진 한 장은 4천만 화소 이하여야 합니다.") from None
    except (UnidentifiedImageError, OSError, ValueError):
        raise WorkerError("INVALID_IMAGE", "사진 파일을 읽을 수 없습니다. 사진을 다시 선택해 주세요.") from None


def crop_scale(width: int, height: int, angle: float) -> float:
    """Largest centered crop with the original aspect, entirely inside rotated bounds.

    Inverse rotation maps the output crop to a parallelogram in the source.
    Both inequalities s*(w*abs(cos)+h*abs(sin)) <= w and
    s*(w*abs(sin)+h*abs(cos)) <= h must hold. This also handles
    panoramas and 45/90 degrees without singularities or alpha-based guesses.
    """
    radians = math.radians(validate_angle(angle))
    cosine, sine = abs(math.cos(radians)), abs(math.sin(radians))
    return min(width / (width * cosine + height * sine), height / (width * sine + height * cosine))


def straighten(image: Image.Image, angle: float) -> Image.Image:
    """Rotate clockwise, crop away introduced corners, keep the same frame size.

    One inverse affine resampling avoids an intermediate expanded raster and
    repeated interpolation. Existing transparent pixels remain transparent.
    Cropping necessarily discards edge content; it never stretches the aspect.
    """
    angle = validate_angle(angle)
    image = image.convert("RGBA")
    if angle == 0:
        return image.copy()
    if abs(angle) == 180:
        return image.transpose(Image.Transpose.ROTATE_180)
    if image.width == image.height and abs(angle) == 90:
        return image.transpose(Image.Transpose.ROTATE_270 if angle > 0 else Image.Transpose.ROTATE_90)
    width, height = image.size
    radians = math.radians(angle)
    cosine, sine = math.cos(radians), math.sin(radians)
    if abs(cosine) < 1e-12:
        cosine = 0.0
    if abs(sine) < 1e-12:
        sine = 0.0
    scale = crop_scale(width, height, angle)
    a, b, d, e = scale * cosine, scale * sine, -scale * sine, scale * cosine
    coefficients = (a, b, width / 2 - a * width / 2 - b * height / 2,
                    d, e, height / 2 - d * width / 2 - e * height / 2)
    # Premultiplied-alpha interpolation prevents dark fringes around transparent edges.
    return image.convert("RGBa").transform(image.size, Image.Transform.AFFINE, coefficients,
        resample=Image.Resampling.BICUBIC).convert("RGBA")


def verify_model(model_dir: Path) -> Path:
    model = model_dir / MODEL_FILENAME
    if not model.is_file():
        raise WorkerError("MODEL_MISSING", "배경 제거 모델이 없습니다. 설치 폴더에서 '모델 준비'를 먼저 실행하거나 u2netp.onnx를 models 폴더에 복사해 주세요. 사진 처리 중에는 모델을 다운로드하지 않습니다.")
    verify_model_file(model)
    return model


def verify_model_file(model: Path) -> None:
    digest = hashlib.md5(usedforsecurity=False)
    with model.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    if digest.hexdigest() != MODEL_MD5:
        raise WorkerError("MODEL_INVALID", "배경 제거 모델 파일이 손상되었거나 다른 모델입니다. 공식 u2netp.onnx로 다시 준비해 주세요.")


def create_background_session(model_dir: Path):
    model = verify_model(model_dir)
    try:
        import onnxruntime as ort
        from rembg.sessions.u2netp import U2netpSession
    except ImportError:
        raise WorkerError("DEPENDENCY_MISSING", "배경 제거 구성 요소가 설치되지 않았습니다. 설치 프로그램의 사진 엔진 준비 단계를 실행해 주세요.") from None

    class OfflineU2netpSession(U2netpSession):
        @classmethod
        def download_models(cls, *args, **kwargs):
            # Never call pooch/new_session: even a corrupt cache must not trigger HTTP.
            return str(model)

    ort.disable_telemetry_events()
    options = ort.SessionOptions()
    options.intra_op_num_threads = min(4, os.cpu_count() or 1)
    options.inter_op_num_threads = 1
    options.log_severity_level = 3
    options.enable_mem_pattern = False
    return OfflineU2netpSession("u2netp", options, providers=["CPUExecutionProvider"])


def remove_background(image: Image.Image, session) -> Image.Image:
    from rembg import remove
    # Request only the inferred mask, preserving RGB and multiplying original alpha.
    mask = remove(image, session=session, only_mask=True, alpha_matting=False,
                  post_process_mask=False)
    if not isinstance(mask, Image.Image) or mask.size != image.size:
        raise WorkerError("PROCESSING_FAILED", "배경 제거 결과 크기가 올바르지 않습니다.")
    result = image.copy()
    result.putalpha(ImageChops.multiply(image.getchannel("A"), mask.convert("L")))
    return result


def process_request(data: object) -> int:
    operation, items, model_dir = parse_request(data)
    staged: list[tuple[Path, Path]] = []
    committed: list[Path] = []
    try:
        session = create_background_session(model_dir) if operation in ("remove_background", "both") else None
        for item in items:
            image = load_image(item.input)
            if session is not None:
                image = remove_background(image, session)
            # Detect the subject once in its original frame. This matches the live
            # preview's cached mask and keeps the outline stable while rotating.
            if operation in ("straighten", "both"):
                image = straighten(image, item.angle)
            fd, name = tempfile.mkstemp(prefix=".labphoto-", suffix=".png", dir=item.output.parent)
            os.close(fd)
            temporary = Path(name)
            staged.append((temporary, item.output))
            image.save(temporary, format="PNG", optimize=False)
            image.close()
        for temporary, destination in staged:
            if os.name == "nt":
                # Windows rename refuses an existing destination, even if a race occurs.
                os.rename(temporary, destination)
            else:
                os.link(temporary, destination)
                temporary.unlink()
            committed.append(destination)
        return len(items)
    except Exception:
        for destination in committed:
            destination.unlink(missing_ok=True)
        raise
    finally:
        for temporary, _ in staged:
            temporary.unlink(missing_ok=True)


def read_request(path: Path) -> object:
    try:
        if path.stat().st_size > MAX_REQUEST_BYTES:
            raise WorkerError("INVALID_REQUEST", "사진 작업 요청이 너무 큽니다.")
        # utf-8-sig also accepts Windows-generated UTF-8 BOM request files.
        return json.loads(path.read_text(encoding="utf-8-sig"),
                          parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
    except (OSError, ValueError, UnicodeError):
        raise WorkerError("INVALID_REQUEST", "사진 작업 요청 파일을 읽을 수 없습니다.") from None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Lab Photo Tools local image worker")
    parser.add_argument("--request", required=True, type=Path)
    args = parser.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        # Third-party Python logs must not corrupt the JSON stdout IPC contract.
        with redirect_stdout(sys.stderr):
            count = process_request(read_request(args.request))
        payload = {"ok": True, "count": count}
        status = 0
    except WorkerError as exc:
        payload = {"ok": False, "error": {"code": exc.code, "message": str(exc)}}
        status = 2
    except MemoryError:
        payload = {"ok": False, "error": {"code": "OUT_OF_MEMORY", "message": "사진을 처리할 메모리가 부족합니다. 사진 크기나 선택 개수를 줄여 주세요."}}
        status = 3
    except Exception:
        payload = {"ok": False, "error": {"code": "PROCESSING_FAILED", "message": "사진 처리에 실패했습니다. 설치 상태와 결과 폴더의 쓰기 권한을 확인해 주세요."}}
        status = 3
    print(json.dumps(payload, ensure_ascii=False))
    return status


if __name__ == "__main__":
    raise SystemExit(main())

