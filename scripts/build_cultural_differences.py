from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = Path("lectures/Cultural_Differences_Lecture.pptx")
CACHE = Path(".image-cache")
CACHE.mkdir(exist_ok=True)
OUT.parent.mkdir(parents=True, exist_ok=True)

# Real editorial photographs; each image has independent fallback sources.
PHOTOS = {
    "cover": [
        "https://www.maharaniweddings.com/wp-content/uploads/2011/06/Indian-wedding-fusion-japanese-1-copy.jpg",
        "https://images.pexels.com/photos/8190802/pexels-photo-8190802.jpeg?auto=compress&cs=tinysrgb&w=1920",
    ],
    "bow": [
        "https://www.cunard.com/content/dam/cunard/marketing-assets/destinations/guides/japan/japanese-bowing-custom-shutterstock-304118024.jpg.image.572.572.low.jpg",
        "https://images.pexels.com/photos/8190802/pexels-photo-8190802.jpeg?auto=compress&cs=tinysrgb&w=1920",
        "https://commons.wikimedia.org/wiki/Special:FilePath/Greeting_when_sitting.jpg",
    ],
    "food": [
        "https://in-sight.symrise.com/fileadmin/_processed_/f/3/csm_multiculturalfoods_f87b945cf0.jpg",
        "https://recipecurrent.com/assets/images/d0261343-1514-4bac-73d5-3fd68a53cb00",
        "https://cf-img-a-in.tosshub.com/sites/visualstory/wp/2025/07/eating-by-hand-2ITG-1751520866321.jpeg?size=%2A%3A900",
    ],
    "nowruz": [
        "https://i0.wp.com/ilearnfarsi.com/wp-content/uploads/2024/06/1c0d5-new-year.jpg?resize=1080%2C675&ssl=1",
        "https://www.happynowruz.com/assets/hero-haftsin-BEfXc3GV.webp",
    ],
    "christmas": [
        "https://images.pexels.com/photos/3303614/pexels-photo-3303614.jpeg?auto=compress&cs=tinysrgb&w=1920",
        "https://www.jufahotels.com/wp-content/uploads/2024/08/weihnachten-familie-christbaum-jufa-hotels-scaled-1620x1080.jpeg.webp",
    ],
    "greeting": [
        "https://cdn-ak.f.st-hatena.com/images/fotolife/k/kentikukun/20250717/20250717215805.jpg",
        "https://images.pexels.com/photos/8190802/pexels-photo-8190802.jpeg?auto=compress&cs=tinysrgb&w=1920",
    ],
}

def get_image(key):
    cached = CACHE / f"{key}.img"
    if cached.exists() and cached.stat().st_size > 5000:
        return cached
    for url in PHOTOS[key]:
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0 CulturalDifferencesSlides/2.0"})
            with urlopen(req, timeout=35) as response:
                data = response.read()
                ctype = response.headers.get("Content-Type", "")
                if len(data) < 5000 or ("image" not in ctype and not data.startswith(b"\xff\xd8") and not data.startswith(b"RIFF") and not data.startswith(b"\x89PNG")):
                    raise ValueError(f"Not a valid image response ({ctype}, {len(data)} bytes)")
                cached.write_bytes(data)
                print(f"Downloaded {key}: {len(data)} bytes")
                return cached
        except Exception as exc:
            print(f"Image source failed for {key}: {url} ({exc})")
    # If a source is temporarily unavailable, reuse a verified photo rather than fail the entire build.
    for fallback_key in ("cover", "food", "greeting", "bow", "nowruz", "christmas"):
        candidate = CACHE / f"{fallback_key}.img"
        if fallback_key != key and candidate.exists() and candidate.stat().st_size > 5000:
            print(f"Using verified {fallback_key} photo as fallback for {key}")
            return candidate
    raise RuntimeError(f"No real photograph could be downloaded for {key}; checked all fallback sources.")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
NAVY = RGBColor(10, 18, 35)
WHITE = RGBColor(255, 255, 255)
GOLD = RGBColor(255, 196, 92)

def add_bg(slide, key, left=0, width=None):
    width = width or prs.slide_width
    slide.shapes.add_picture(str(get_image(key)), left, 0, width=width, height=prs.slide_height)

def overlay(slide, left=0, width=None, transparency=35):
    width = width or prs.slide_width
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, 0, width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.fill.transparency = transparency
    shape.line.fill.background()

def text(slide, value, x, y, w, h, size, color=WHITE, bold=False, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.04)
    tf.margin_right = Inches(.04)
    tf.margin_top = Inches(.02)
    tf.margin_bottom = Inches(.02)
    p = tf.paragraphs[0]
    p.text = value
    p.font.name = "Aptos Display"
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    if align is not None:
        p.alignment = align
    return box

def pill(slide, value, x=.75, y=.55):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.65), Inches(.42))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = value.upper()
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Aptos"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY

def footer(slide, n, credit):
    text(slide, f"{n:02d}   •   {credit}", .75, 7.12, 11.9, .22, 9, WHITE)

# 1 — Cover
s = prs.slides.add_slide(blank)
add_bg(s, "cover")
overlay(s, transparency=28)
pill(s, "English presentation")
text(s, "CULTURAL\nDIFFERENCES", .8, 2.0, 11.8, 2.25, 48, WHITE, True)
text(s, "One world. Many ways of life.", .9, 4.55, 10.8, .55, 24, GOLD, True)
text(s, "FOOD  •  GREETINGS  •  TRADITIONS  •  RESPECT", .9, 6.45, 11.5, .35, 13, WHITE, True)
footer(s, 1, "Photo: Maharani Weddings")

# 2 — What is culture?
s = prs.slides.add_slide(blank)
add_bg(s, "food")
overlay(s, transparency=30)
pill(s, "Everyday life")
text(s, "WHAT MAKES A CULTURE?", .8, 1.55, 11.8, 1.0, 39, WHITE, True)
text(s, "FOOD", .9, 3.35, 2.3, .6, 25, GOLD, True)
text(s, "LANGUAGE", 3.6, 3.35, 3.0, .6, 25, WHITE, True)
text(s, "VALUES", 7.0, 3.35, 2.3, .6, 25, GOLD, True)
text(s, "TRADITIONS", 9.55, 3.35, 3.0, .6, 25, WHITE, True)
text(s, "Culture shapes how we live and connect.", .9, 5.2, 10.5, .6, 21, WHITE)
footer(s, 2, "Photo: Symrise In-Sight")

# 3 — Greetings
s = prs.slides.add_slide(blank)
add_bg(s, "bow")
overlay(s, transparency=27)
pill(s, "Different greetings")
text(s, "A BOW OR A\nHANDSHAKE?", .8, 1.65, 8.7, 1.9, 43, WHITE, True)
text(s, "JAPAN  •  BOWING", .9, 4.25, 5.8, .45, 23, GOLD, True)
text(s, "Many other places  •  Handshakes", .9, 4.9, 7.5, .45, 19, WHITE)
text(s, "The meaning is respect.", .9, 5.8, 7.5, .45, 20, WHITE, True)
footer(s, 3, "Photo: Cunard / Pexels")

# 4 — Food
s = prs.slides.add_slide(blank)
add_bg(s, "food")
overlay(s, transparency=24)
pill(s, "Food & habits")
text(s, "TASTE THE\nDIFFERENCE", .8, 1.6, 8.2, 1.8, 45, WHITE, True)
text(s, "Recipes tell stories.", .9, 4.25, 6.0, .55, 25, GOLD, True)
text(s, "Eating habits can differ, too.", .9, 4.95, 7.8, .55, 20, WHITE)
footer(s, 4, "Photo: Symrise In-Sight")

# 5 — Celebrations
s = prs.slides.add_slide(blank)
add_bg(s, "nowruz", 0, Inches(6.666))
add_bg(s, "christmas", Inches(6.666), Inches(6.667))
for x in (0, Inches(6.666)):
    overlay(s, x, Inches(6.667), 48)
pill(s, "Celebrations", .72, .55)
text(s, "NOWRUZ", .75, 2.25, 5.7, .7, 34, WHITE, True)
text(s, "Spring  •  Haft-Seen  •  New beginnings", .8, 3.15, 5.65, .85, 17, GOLD, True)
text(s, "CHRISTMAS", 7.35, 2.25, 5.4, .7, 34, WHITE, True)
text(s, "December 25  •  Trees  •  Giving", 7.4, 3.15, 5.4, .85, 17, GOLD, True)
text(s, "Different celebrations, shared joy.", 2.0, 6.35, 9.3, .45, 20, WHITE, True, PP_ALIGN.CENTER)
footer(s, 5, "Photos: iLearn Farsi / Pexels")

# 6 — Everyday communication
s = prs.slides.add_slide(blank)
add_bg(s, "greeting")
overlay(s, transparency=27)
pill(s, "Daily communication")
text(s, "SMALL HABITS.\nBIG MEANINGS.", .8, 1.6, 9.5, 1.8, 43, WHITE, True)
text(s, "TIME       PERSONAL SPACE", .9, 4.15, 9.7, .5, 22, GOLD, True)
text(s, "GESTURES       BODY LANGUAGE", .9, 4.85, 10.5, .5, 22, WHITE, True)
text(s, "Observe first. Never assume.", .9, 5.8, 8.0, .45, 19, WHITE)
footer(s, 6, "Photo: Singapore Info")

# 7 — Closing
s = prs.slides.add_slide(blank)
add_bg(s, "cover")
overlay(s, transparency=30)
pill(s, "The message")
text(s, "DIFFERENT DOESN'T\nMEAN WRONG.", .8, 1.65, 11.8, 1.9, 42, WHITE, True)
text(s, "RESPECT  •  CURIOSITY  •  UNDERSTANDING", .9, 4.2, 11.4, .55, 19, GOLD, True)
text(s, "THANK YOU!", .9, 5.35, 8.0, .75, 34, WHITE, True)
footer(s, 7, "Photo: Maharani Weddings")

prs.core_properties.title = "Cultural Differences"
prs.core_properties.subject = "English class presentation"
prs.core_properties.author = "MMD Studio"
prs.save(OUT)
print(f"Created {OUT} ({OUT.stat().st_size} bytes, {len(prs.slides)} slides)")
