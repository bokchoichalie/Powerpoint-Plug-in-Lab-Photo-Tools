"""Check a relocated embedded runtime without relying on an installed Python."""
import json
from pathlib import Path
import socket
import subprocess
import sys

app = Path(sys.argv[1]).resolve()
runtime = app / "runtime"
assert sys.version_info[:2] == (3, 12)
assert sys.flags.isolated and sys.flags.ignore_environment
assert Path(sys.executable).is_relative_to(runtime)
assert not (runtime / "pyvenv.cfg").exists()
import PIL
import numpy
import onnxruntime
import rembg
for module in (PIL, numpy, onnxruntime, rembg):
    assert Path(module.__file__).resolve().is_relative_to(runtime), module.__file__
from PIL import Image, ImageDraw
import worker

# Inference must use the installed model, with no network fallback.
def offline(*args, **kwargs):
    raise AssertionError("Image processing attempted a network connection")
socket.create_connection = offline
socket.socket.connect = offline
image = Image.new("RGBA", (160, 120), "white")
ImageDraw.Draw(image).ellipse((45, 15, 115, 105), fill=(180, 30, 30, 255))
rotated = worker.straighten(image, 12.5)
assert rotated.size == image.size
session = worker.create_background_session(app / "models")
removed = worker.remove_background(image, session)
assert removed.size == image.size and removed.mode == "RGBA"
assert removed.getchannel("A").getextrema()[0] < 200
assert removed.getpixel((80, 60))[3] > 100
image.save(app / "input.png")
for operation in ("straighten", "remove_background"):
    output = app / (operation + ".png")
    output.unlink(missing_ok=True)
    request = app / (operation + ".json")
    request.write_text(json.dumps({"operation": operation, "model_dir": str(app / "models"),
        "items": [{"input": str(app / "input.png"), "output": str(output), "angle": 12.5}]}), encoding="utf-8")
    result = subprocess.run([sys.executable, str(app / "engine" / "worker.py"), "--request", str(request)],
        capture_output=True, text=True, encoding="utf-8", timeout=90)
    assert result.returncode == 0, (result.returncode, result.stdout, result.stderr)
    assert json.loads(result.stdout)["ok"]
    assert Image.open(output).size == image.size
print("PASS: isolated embedded Python, relocated dependencies, offline inference and both worker CLI operations")
