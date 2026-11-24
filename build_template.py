"""Recreate the macro-free Excel template. Python 3, standard library only."""
from base64 import b64decode
from pathlib import Path

data = (Path(__file__).parent / "template.xlsx.b64").read_text(encoding="ascii")
target = Path(__file__).parent / "Noise_R_Us_Template.xlsx"
if target.exists():
    raise SystemExit("Template already exists; refusing to overwrite it.")
target.write_bytes(b64decode(data, validate=True))
print("Created", target)
