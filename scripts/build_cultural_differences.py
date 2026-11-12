from pathlib import Path
from urllib.request import Request, urlopen
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR

OUT = Path("lectures/Cultural_Differences_Lecture.pptx")
CACHE = Path(".image-cache")
CACHE.mkdir(exist_ok=True)
OUT.parent.mkdir(parents=True, exist_ok=True)

# Real editorial / photo-hosted images (not AI-generated). Each has a fallback.
PHOTOS = {
 "cover": [
  "https://inro.pdn.ac.lk/assets/images/testimonial/Dilki/10.jpg",
  "https://www.maharaniweddings.com/wp-content/uploads/2011/06/Indian-wedding-fusion-japanese-1-copy.jpg"
 ],
 "bow": [
  "https://www.newsweekjapan.jp/stories/assets_c/2019/03/almomen190312-japan-thumb-720x483-154264.jpg"
 ],
 "food": [
  "https://cf-img-a-in.tosshub.com/sites/visualstory/wp/2025/07/eating-by-hand-2ITG-1751520866321.jpeg?size=%2A%3A900",
  "https://recipecurrent.com/assets/images/d0261343-1514-4bac-73d5-3fd68a53cb00"
 ],
 "nowruz": ["https://www.happynowruz.com/assets/hero-haftsin-BEfXc3GV.webp"],
 "christmas": [
  "https://www.jufahotels.com/wp-content/uploads/2024/08/weihnachten-familie-christbaum-jufa-hotels-scaled-1620x1080.jpeg.webp",
  "https://images.pexels.com/photos/3303614/pexels-photo-3303614.jpeg?auto=compress&cs=tinysrgb&w=1920"
 ],
 "greeting": [
  "https://cdn-ak.f.st-hatena.com/images/fotolife/k/kentikukun/20250717/20250717215805.jpg"
 ],
}

def get_image(key):
    path = CACHE / f"{key}.img"
    if path.exists() and path.stat().st_size > 5000:
        return path
    for url in PHOTOS[key]:
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0 (compatible; CulturalDifferencesSlides/1.0)"})
            with urlopen(req, timeout=25) as response:
                data = response.read()
            if len(data) < 5000:
                continue
            path.write_bytes(data)
            return path
        except Exception as exc:
            print(f"Image source failed for {key}: {url} ({exc})")
    raise RuntimeError(f"No usable real photo could be downloaded for: {key}")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
NAVY = RGBColor(10, 18, 35)
WHITE = RGBColor(255,255,255)
GOLD = RGBColor(255,196,92)
MINT = RGBColor(130,231,205)

def add_bg(slide, key):
    slide.shapes.add_picture(str(get_image(key)), 0, 0, width=prs.slide_width, height=prs.slide_height)
    overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    overlay.fill.solid(); overlay.fill.fore_color.rgb = NAVY; overlay.fill.transparency = 35
    overlay.line.fill.background()

def textbox(slide, text, x, y, w, h, size, color=WHITE, bold=False, align=None):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]; p.text = text
    p.font.name = "Aptos Display"; p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = color
    if align is not None: p.alignment = align
    return shape

def tag(slide, text, x=0.72, y=0.55):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.5), Inches(0.42))
    s.fill.solid(); s.fill.fore_color.rgb = GOLD; s.line.fill.background()
    tf=s.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0]; p.text=text.upper(); p.font.name="Aptos"; p.font.size=Pt(11); p.font.bold=True; p.font.color.rgb=NAVY; p.alignment=PP_ALIGN.CENTER

def footer(slide, n, credit):
    textbox(slide, f"{n:02d}  •  {credit}", 0.75, 7.08, 11.8, 0.22, 9, WHITE)

# 1 — Cover
s=prs.slides.add_slide(blank); add_bg(s,"cover")
tag(s,"English presentation")
textbox(s,"CULTURAL\nDIFFERENCES",0.8,2.0,8.9,2.35,48,WHITE,True)
textbox(s,"One world. Many ways of life.",0.88,4.55,7.5,0.6,23,GOLD,True)
textbox(s,"FOOD  •  GREETINGS  •  TRADITIONS  •  RESPECT",0.9,6.55,10.5,0.35,12,WHITE,True)
footer(s,1,"Photo: University of Peradeniya / Maharani Weddings")

# 2 — Culture
s=prs.slides.add_slide(blank); add_bg(s,"food"); tag(s,"What makes a culture?")
textbox(s,"CULTURE IS\nEVERYDAY LIFE",0.8,1.65,6.8,1.7,40,WHITE,True)
textbox(s,"FOOD     LANGUAGE\nVALUES   TRADITIONS",0.9,4.1,5.8,1.15,23,GOLD,True)
textbox(s,"It shapes how we live and connect.",0.9,5.65,6.5,0.55,18,WHITE)
footer(s,2,"Photo: India Today / RecipeCurrent")

# 3 — Greetings
s=prs.slides.add_slide(blank); add_bg(s,"bow"); tag(s,"Different greetings")
textbox(s,"A BOW OR\nA HANDSHAKE?",0.8,1.7,7.5,1.9,43,WHITE,True)
textbox(s,"Japan  •  Bowing\nMany places  •  Handshakes",0.9,4.55,7.0,1.05,22,GOLD,True)
textbox(s,"The meaning is respect.",0.9,5.85,6.2,0.45,18,WHITE)
footer(s,3,"Photo: Newsweek Japan")

# 4 — Food
s=prs.slides.add_slide(blank); add_bg(s,"food"); tag(s,"Food & habits")
textbox(s,"TASTE THE\nDIFFERENCE",0.8,1.55,6.5,1.8,43,WHITE,True)
textbox(s,"Recipes tell stories.\nEating habits can differ, too.",0.9,4.15,6.7,1.05,22,GOLD,True)
footer(s,4,"Photo: India Today / RecipeCurrent")

# 5 — Celebrations split
s=prs.slides.add_slide(blank)
s.shapes.add_picture(str(get_image("nowruz")),0,0,width=Inches(6.67),height=prs.slide_height)
s.shapes.add_picture(str(get_image("christmas")),Inches(6.67),0,width=Inches(6.67),height=prs.slide_height)
for x in (0,6.67):
    ov=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),0,Inches(6.67),prs.slide_height)
    ov.fill.solid(); ov.fill.fore_color.rgb=NAVY; ov.fill.transparency=45; ov.line.fill.background()
tag(s,"Celebrations",0.72,0.55)
textbox(s,"NOWRUZ",0.75,2.35,5.3,0.8,36,WHITE,True)
textbox(s,"Spring • Haft-Seen • New beginnings",0.8,3.35,5.5,1.0,17,GOLD,True)
textbox(s,"CHRISTMAS",7.35,2.35,5.3,0.8,36,WHITE,True)
textbox(s,"December 25 • Trees • Giving",7.4,3.35,5.2,1.0,17,GOLD,True)
textbox(s,"Different celebrations, shared joy.",2.6,6.45,8.5,0.5,20,WHITE,True,PP_ALIGN.CENTER)
footer(s,5,"Photos: HappyNowruz / JUFA Hotels")

# 6 — Everyday communication
s=prs.slides.add_slide(blank); add_bg(s,"greeting"); tag(s,"Daily communication")
textbox(s,"SMALL HABITS.\nBIG MEANINGS.",0.8,1.55,7.8,1.8,42,WHITE,True)
textbox(s,"TIME     PERSONAL SPACE\nGESTURES     BODY LANGUAGE",0.9,4.2,8.5,1.1,21,GOLD,True)
textbox(s,"Observe first. Never assume.",0.9,5.75,6.5,0.5,18,WHITE)
footer(s,6,"Photo: Singapore Info")

# 7 — Closing
s=prs.slides.add_slide(blank); add_bg(s,"cover")
tag(s,"The message")
textbox(s,"DIFFERENT DOESN'T\nMEAN WRONG.",0.8,1.55,10.8,1.9,42,WHITE,True)
textbox(s,"RESPECT  •  CURIOSITY  •  UNDERSTANDING",0.9,4.25,10.5,0.55,20,GOLD,True)
textbox(s,"THANK YOU!",0.9,5.55,5.5,0.8,34,WHITE,True)
footer(s,7,"Photo: University of Peradeniya / Maharani Weddings")

prs.core_properties.title = "Cultural Differences"
prs.core_properties.subject = "English class presentation"
prs.core_properties.author = "MMD Studio"
prs.save(OUT)
print(f"Created {OUT} ({OUT.stat().st_size} bytes, {len(prs.slides)} slides)")
