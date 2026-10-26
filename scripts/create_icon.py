"""Create the desktop icons with Pillow (build-time only)."""
from pathlib import Path
from PIL import Image, ImageDraw

root = Path(__file__).resolve().parents[1]
assets = root / "assets"
assets.mkdir(exist_ok=True)

scale = 4
im = Image.new("RGBA", (1024 * scale, 1024 * scale), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
S = lambda n: n * scale
d.rounded_rectangle((S(40), S(40), S(984), S(984)), radius=S(220), fill="#126F69")
d.rounded_rectangle((S(222), S(150), S(803), S(850)), radius=S(54), fill="#FFFFFF")
d.rounded_rectangle((S(278), S(250), S(747), S(302)), radius=S(10), fill="#C15345")
d.rounded_rectangle((S(310), S(392), S(714), S(418)), radius=S(8), fill="#CED8D4")
d.rounded_rectangle((S(310), S(473), S(690), S(499)), radius=S(8), fill="#CED8D4")
d.rounded_rectangle((S(310), S(554), S(714), S(580)), radius=S(8), fill="#CED8D4")
d.rounded_rectangle((S(310), S(635), S(602), S(661)), radius=S(8), fill="#CED8D4")
d.ellipse((S(565), S(625), S(816), S(876)), fill="#C15345")
d.ellipse((S(584), S(644), S(797), S(857)), outline="#FFFFFF", width=S(9))
d.line((S(640), S(745), S(680), S(785), S(749), S(706)), fill="#FFFFFF", width=S(21), joint="curve")
im = im.resize((1024, 1024), Image.Resampling.LANCZOS)
im.save(assets / "app.png")
im.save(assets / "app.ico", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
im.save(assets / "app.icns")
print("Created app.png, app.ico, app.icns")
