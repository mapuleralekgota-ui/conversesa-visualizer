from pathlib import Path
import base64


BASE = Path(__file__).resolve().parent
HTML = BASE / "index.html"
ASSETS = {
    "shoes": "real-shoes.jpg",
    "crowd": "real-crowd.jpg",
    "street": "real-street.jpg",
}


html = HTML.read_text(encoding="utf-8")

variables = []
for key, name in ASSETS.items():
    image_data = base64.b64encode((BASE / "real-assets" / name).read_bytes()).decode("ascii")
    variables.append(f'      --image-{key}: url("data:image/jpeg;base64,{image_data}");')

marker = "      --shadow: 0 28px 80px rgba(0,0,0,.38);"
if "--image-shoes:" not in html:
    html = html.replace(marker, marker + "\n" + "\n".join(variables))

for key, name in ASSETS.items():
    html = html.replace(f"url('real-assets/{name}')", f"var(--image-{key})")

HTML.write_text(html, encoding="utf-8")
print(f"Embedded images into {HTML}")
