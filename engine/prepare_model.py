"""Explicit one-time model preparation; no photos are read by this command."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
import tempfile
from urllib.request import urlopen

from worker import MODEL_FILENAME, WorkerError, absolute_path, verify_model, verify_model_file

MODEL_URL = "https://github.com/danielgatis/rembg/releases/download/v0.0.0/u2netp.onnx"
MAX_MODEL_BYTES = 10 * 1024 * 1024


def prepare(model_dir: Path) -> bool:
    if (model_dir / MODEL_FILENAME).exists():
        verify_model(model_dir)
        return False
    model_dir.mkdir(parents=True, exist_ok=True)
    # Stay in the destination folder so Windows file ACLs come from that folder,
    # not a private TemporaryDirectory whose restricted ACL survives rename.
    fd, temporary = tempfile.mkstemp(prefix=".model-", suffix=".onnx", dir=model_dir)
    os.close(fd)
    pending = Path(temporary)
    try:
        with urlopen(MODEL_URL, timeout=60) as response, pending.open("wb") as output:
            size = 0
            while chunk := response.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_MODEL_BYTES:
                    raise WorkerError("MODEL_INVALID", "모델 다운로드 크기가 예상 범위를 초과했습니다.")
                output.write(chunk)
        verify_model_file(pending)
        destination = model_dir / MODEL_FILENAME
        if os.name == "nt":
            os.rename(pending, destination)
        else:
            os.link(pending, destination)
            pending.unlink()
    finally:
        pending.unlink(missing_ok=True)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Download the official u2netp model once (internet required; no photo upload).")
    parser.add_argument("--model-dir", required=True)
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        downloaded = prepare(absolute_path(args.model_dir, "모델 폴더"))
        print(json.dumps({"ok": True, "downloaded": downloaded}, ensure_ascii=False))
        return 0
    except WorkerError as exc:
        print(json.dumps({"ok": False, "error": {"code": exc.code, "message": str(exc)}}, ensure_ascii=False))
    except Exception:
        print(json.dumps({"ok": False, "error": {"code": "MODEL_PREPARE_FAILED", "message": "모델 준비에 실패했습니다. 인터넷 연결과 모델 폴더 권한을 확인해 주세요. 다른 PC에서 받은 공식 u2netp.onnx를 복사해도 됩니다."}}, ensure_ascii=False))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())

