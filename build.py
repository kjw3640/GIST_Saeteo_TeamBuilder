"""Build a self-contained offline page. Deployment contains no source workbooks."""
import base64
import hashlib
import zipfile
from pathlib import Path

root = Path(__file__).resolve().parent
src = root / "src"
js = "\n".join((src / f).read_text() for f in ("book.js", "input.js", "solver.js", "ui.js"))
css = (src / "style.css").read_text()

def digest(value):
    return "'sha256-" + base64.b64encode(hashlib.sha256(value.encode()).digest()).decode() + "'"

policy = "; ".join(["default-src 'none'", "script-src " + digest(js), "style-src " + digest(css),
                    "img-src data:", "worker-src blob:", "connect-src 'none'", "object-src 'none'",
                    "base-uri 'none'", "form-action 'none'"])
out = root / "dist"
out.mkdir(exist_ok=True)
html = (src / "page.html").read_text().replace("__POLICY__", policy).replace("__STYLE__", css).replace("__SCRIPT__", js)
(out / "index.html").write_text(html)
(out / "_headers").write_text("/*\n  Content-Security-Policy: " + policy + "; frame-ancestors 'none'\n  X-Frame-Options: DENY\n  X-Content-Type-Options: nosniff\n  Referrer-Policy: no-referrer\n  Cache-Control: no-store\n  Permissions-Policy: camera=(), microphone=(), geolocation=()\n")

with zipfile.ZipFile(out / "saeteo-cli.zip", "w", zipfile.ZIP_DEFLATED) as archive:
    for filename in ("main.py", "requirements.txt", "사용법.txt"):
        item = zipfile.ZipInfo(filename, (2026, 2, 21, 0, 0, 0))
        archive.writestr(item, (root / filename).read_bytes(), compress_type=zipfile.ZIP_DEFLATED)
