import os
import qrcode
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Setup Paths & Dimensions
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
ASSETS_DIR = os.path.join(PROJECT_ROOT, "assets")
GEN_DIR = os.path.join(SCRIPT_DIR, "gen_assets")
os.makedirs(GEN_DIR, exist_ok=True)

OUTPUT_PPTX = os.path.join(SCRIPT_DIR, "BaySec_Week4_Kickoff.pptx")

# 16:9 Widescreen dimensions
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Color Palette
BG_COLOR = RGBColor(11, 15, 25)         # #0B0F19 - Deep cyber navy
CARD_BG = RGBColor(17, 24, 39)          # #111827 - Dark slate card
CARD_BORDER = RGBColor(31, 41, 61)      # #1F293D
CYAN = RGBColor(0, 240, 255)            # #00F0FF - Electric Cyan
BLUE_ACCENT = RGBColor(56, 189, 248)    # #38BDF8 - Sky Blue
RED_ACCENT = RGBColor(239, 68, 68)      # #EF4444 - Crimson Red
PURPLE_ACCENT = RGBColor(168, 85, 247)  # #A855F7 - AI Security Purple
GREEN_ACCENT = RGBColor(74, 222, 128)   # #4ADE80 - Terminal Green
WHITE = RGBColor(248, 250, 252)         # #F8FAFC
SLATE_LIGHT = RGBColor(203, 213, 225)   # #CBD5E1
SLATE_MUTED = RGBColor(148, 163, 184)   # #94A3B8
AMBER = RGBColor(245, 158, 11)          # #F59E0B

FONT_CODE = "Consolas"
FONT_BODY = "Segoe UI"
FONT_HEADING = "Segoe UI"

# ---------------------------------------------------------------------------
# QR Code Generator
# ---------------------------------------------------------------------------
def make_qr(data: str, filename: str, fill_color="black", back_color="white") -> str:
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill_color, back_color=back_color)
    path = os.path.join(GEN_DIR, filename)
    img.save(path)
    return path

QR_ATTENDANCE = make_qr("https://forms.gle/b5oD4hmCSG41ZcLp7", "qr_attendance.png")
QR_WHATSAPP = os.path.join(ASSETS_DIR, "qr.png")
QR_CHALLENGE = make_qr("https://MynameisKoi.github.io/baysec/week4/", "qr_challenge.png")

BLUE_LOGO = os.path.join(ASSETS_DIR, "baysec-blue.png")
RED_LOGO = os.path.join(ASSETS_DIR, "baysec-red.png")

# ---------------------------------------------------------------------------
# PPTX Helper Functions
# ---------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT
blank_layout = prs.slide_layouts[6]

def create_slide():
    slide = prs.slides.add_slide(blank_layout)
    # Background full rectangle
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_COLOR
    bg.line.fill.background()
    return slide

def add_header(slide, title_text, category_text=None, top_offset=Inches(0.4)):
    if category_text:
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            top_offset,
            Inches(2.5),
            Inches(0.35)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(15, 23, 42)
        badge.line.color.rgb = CYAN
        badge.line.width = Pt(1)
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = category_text.upper()
        p.font.name = FONT_CODE
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = CYAN
        p.alignment = PP_ALIGN.CENTER
        top_offset += Inches(0.42)

    tb = slide.shapes.add_textbox(Inches(0.8), top_offset, Inches(11.733), Inches(0.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_HEADING
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE
    return top_offset + Inches(0.9)

def add_speaker_notes(slide, notes_text: str):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text

# ===========================================================================
# SLIDE 1: Title & Check-In
# ===========================================================================
slide1 = create_slide()

# Top Tag Badge
tag = slide1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(5.0),
    Inches(0.6),
    Inches(3.333),
    Inches(0.38)
)
tag.fill.solid()
tag.fill.fore_color.rgb = RGBColor(15, 23, 42)
tag.line.color.rgb = CYAN
tag.line.width = Pt(1.5)
tf = tag.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "KICKOFF MEETING | FALL 2026"
p.font.name = FONT_CODE
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = CYAN
p.alignment = PP_ALIGN.CENTER

# Centered Title
title_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.15), Inches(11.333), Inches(1.3))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "BaySec @ SFBU"
p.font.name = FONT_HEADING
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Subtitle
sub_box = slide1.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.333), Inches(0.6))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.text = "Securing the Bay, One Byte at a Time"
p.font.name = FONT_BODY
p.font.size = Pt(20)
p.font.color.rgb = BLUE_ACCENT
p.alignment = PP_ALIGN.CENTER

# Centered Dual Logos
if os.path.exists(BLUE_LOGO) and os.path.exists(RED_LOGO):
    logo_w = Inches(2.2)
    logo_h = Inches(2.2)
    center_x = SLIDE_WIDTH / 2
    slide1.shapes.add_picture(BLUE_LOGO, center_x - Inches(2.35), Inches(3.0), logo_w, logo_h)
    slide1.shapes.add_picture(RED_LOGO, center_x + Inches(0.15), Inches(3.0), logo_w, logo_h)

# Welcome text beneath logos
wel_box = slide1.shapes.add_textbox(Inches(2.5), Inches(5.35), Inches(8.333), Inches(0.8))
tf = wel_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Welcome! Please scan the QR codes to check in and join our active channels while we set up."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = SLATE_LIGHT
p.alignment = PP_ALIGN.CENTER

# QR Code 1 (Bottom-Left)
card_qr1 = slide1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    Inches(4.9),
    Inches(2.3),
    Inches(2.2)
)
card_qr1.fill.solid()
card_qr1.fill.fore_color.rgb = CARD_BG
card_qr1.line.color.rgb = CYAN
card_qr1.line.width = Pt(1)
slide1.shapes.add_picture(QR_ATTENDANCE, Inches(1.15), Inches(5.05), Inches(1.6), Inches(1.6))
lbl1 = slide1.shapes.add_textbox(Inches(0.8), Inches(6.65), Inches(2.3), Inches(0.4))
p1 = lbl1.text_frame.paragraphs[0]
p1.text = "Scan for Attendance"
p1.font.name = FONT_BODY
p1.font.size = Pt(11)
p1.font.bold = True
p1.font.color.rgb = CYAN
p1.alignment = PP_ALIGN.CENTER

# QR Code 2 (Bottom-Right)
card_qr2 = slide1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(10.233),
    Inches(4.9),
    Inches(2.3),
    Inches(2.2)
)
card_qr2.fill.solid()
card_qr2.fill.fore_color.rgb = CARD_BG
card_qr2.line.color.rgb = GREEN_ACCENT
card_qr2.line.width = Pt(1)
slide1.shapes.add_picture(QR_WHATSAPP, Inches(10.583), Inches(5.05), Inches(1.6), Inches(1.6))
lbl2 = slide1.shapes.add_textbox(Inches(10.233), Inches(6.65), Inches(2.3), Inches(0.4))
p2 = lbl2.text_frame.paragraphs[0]
p2.text = "Join WhatsApp Community"
p2.font.name = FONT_BODY
p2.font.size = Pt(11)
p2.font.bold = True
p2.font.color.rgb = GREEN_ACCENT
p2.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide1,
    "Welcome everyone to the official kickoff of BaySec @ SFBU! While everyone gets seated, "
    "take a moment to scan the QR codes on screen to mark your attendance and get plugged into our "
    "official WhatsApp group. That’s where we’ll drop lab resources, room updates, and CTF announcements."
)

# ===========================================================================
# SLIDE 2: Welcome & The Teaser Solved
# ===========================================================================
slide2 = create_slide()
content_top = add_header(slide2, "Welcome to BaySec", "Introduction & Reveal")

col_w = Inches(5.6)
col_h = Inches(5.1)

# Left Column: Mission & Leadership
left_card = slide2.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    content_top,
    col_w,
    col_h
)
left_card.fill.solid()
left_card.fill.fore_color.rgb = CARD_BG
left_card.line.color.rgb = CYAN
left_card.line.width = Pt(1.5)

tf_l = left_card.text_frame
tf_l.word_wrap = True
tf_l.margin_left = Inches(0.4)
tf_l.margin_right = Inches(0.4)
tf_l.margin_top = Inches(0.4)

p = tf_l.paragraphs[0]
p.text = "🛡️ OUR MISSION"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf_l.add_paragraph()
p.text = "Bridging academic computer science with real-world ethical hacking, active system defense, and modern AI security."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = WHITE
p.space_before = Pt(10)
p.space_after = Pt(24)

p = tf_l.add_paragraph()
p.text = "👥 LEADERSHIP"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = BLUE_ACCENT

p = tf_l.add_paragraph()
p.text = "• Khoi Duong — Club President"
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_before = Pt(8)

p = tf_l.add_paragraph()
p.text = "   mduong@student.sfbu.edu"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.color.rgb = SLATE_MUTED

p = tf_l.add_paragraph()
p.text = "• Joed Prudente — Faculty Advisor"
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_before = Pt(12)

p = tf_l.add_paragraph()
p.text = "   joed.prudente@sfbu.edu"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.color.rgb = SLATE_MUTED

# Right Column: The Teaser Posters Decoded
right_card = slide2.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(6.933),
    content_top,
    col_w,
    col_h
)
right_card.fill.solid()
right_card.fill.fore_color.rgb = CARD_BG
right_card.line.color.rgb = RED_ACCENT
right_card.line.width = Pt(1.5)

tf_r = right_card.text_frame
tf_r.word_wrap = True
tf_r.margin_left = Inches(0.4)
tf_r.margin_right = Inches(0.4)
tf_r.margin_top = Inches(0.4)

p = tf_r.paragraphs[0]
p.text = "🔓 THE TEASER POSTERS DECODED"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = RED_ACCENT

p = tf_r.add_paragraph()
p.text = "Did you crack the secret flyers across campus?"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = SLATE_LIGHT
p.space_before = Pt(6)
p.space_after = Pt(14)

# Cipher Box 1: Hex
p = tf_r.add_paragraph()
p.text = "CIPHER 1: HEXADECIMAL STREAM"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf_r.add_paragraph()
p.text = "RAW:  42 61 79 53 65 63 20 69 73 20 68 65 72 65"
p.font.name = FONT_CODE
p.font.size = Pt(11)
p.font.color.rgb = SLATE_LIGHT

p = tf_r.add_paragraph()
p.text = "SOLVED ➔ \"BaySec is here\""
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.space_after = Pt(16)

# Cipher Box 2: ROT-13
p = tf_r.add_paragraph()
p.text = "CIPHER 2: ROT-13 SUBSTITUTION"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf_r.add_paragraph()
p.text = "RAW:  Onlfrp vf urer"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.color.rgb = SLATE_LIGHT

p = tf_r.add_paragraph()
p.text = "SOLVED ➔ \"BaySec is here\""
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.space_after = Pt(14)

p = tf_r.add_paragraph()
p.text = "✓ Both paths converged. Welcome to the club."
p.font.name = FONT_BODY
p.font.size = Pt(12)
p.font.italic = True
p.font.color.rgb = CYAN

add_speaker_notes(
    slide2,
    "If you saw our flyers around campus, you might have tackled the ciphers. "
    "Both the hexadecimal sequence and the ROT-13 text translated to 'BaySec is here.' "
    "Today, we’re officially here. Our goal is simple: take the concepts you study in class "
    "and apply them directly to live defense, penetration testing, and security tooling."
)

# ===========================================================================
# SLIDE 3: The Three Pillars
# ===========================================================================
slide3 = create_slide()
content_top = add_header(slide3, "Choose Your Path (Or Explore All Three)", "Curriculum Tracks")

card_w = Inches(3.64)
card_h = Inches(4.3)
gap = Inches(0.4)
left_base = Inches(0.8)

pillars = [
    {
        "title": "BLUE TEAM",
        "subtitle": "Active Defense & Triage",
        "color": CYAN,
        "items": [
            "Threat detection & incident triage",
            "Network traffic analysis (Wireshark)",
            "SIEM monitoring & defensive posture",
            "Log auditing & host forensics"
        ]
    },
    {
        "title": "RED TEAM",
        "subtitle": "Offense & Penetration Testing",
        "color": RED_ACCENT,
        "items": [
            "Web application security & OWASP Top 10",
            "Ethical exploitation & penetration testing",
            "Competitive Jeopardy-style CTF challenges",
            "Privilege escalation & binary analysis"
        ]
    },
    {
        "title": "EMERGING TECH",
        "subtitle": "AI Security & Modern Frontiers",
        "color": PURPLE_ACCENT,
        "items": [
            "LLM red-teaming & prompt injection",
            "Synthetic media & deepfake detection",
            "Defending autonomous agent workflows",
            "Adversarial ML attack mitigations"
        ]
    }
]

for i, p_info in enumerate(pillars):
    card = slide3.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left_base + i * (card_w + gap),
        content_top,
        card_w,
        card_h
    )
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = p_info["color"]
    card.line.width = Pt(2)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.35)

    p = tf.paragraphs[0]
    p.text = p_info["title"]
    p.font.name = FONT_CODE
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = p_info["color"]

    p = tf.add_paragraph()
    p.text = p_info["subtitle"]
    p.font.name = FONT_BODY
    p.font.size = Pt(12)
    p.font.color.rgb = SLATE_MUTED
    p.space_after = Pt(18)

    for item in p_info["items"]:
        p = tf.add_paragraph()
        p.text = f"▸  {item}"
        p.font.name = FONT_BODY
        p.font.size = Pt(13)
        p.font.color.rgb = WHITE
        p.space_before = Pt(8)

# Banner Callout at bottom
banner = slide3.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    Inches(6.25),
    Inches(11.733),
    Inches(0.65)
)
banner.fill.solid()
banner.fill.fore_color.rgb = RGBColor(15, 23, 42)
banner.line.color.rgb = BLUE_ACCENT
banner.line.width = Pt(1)
tf_b = banner.text_frame
tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf_b.paragraphs[0]
p.text = "⚡ Zero prior experience required. You do not need to lock yourself into a single path."
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide3,
    "You don’t have to pick just one discipline. Whether you want to learn how attackers "
    "break into web apps, how SOC analysts hunt threats in network logs, or how prompt injection "
    "cracks AI systems, you can move between tracks based on what excites you."
)

# ===========================================================================
# SLIDE 4: Semester Roadmap Highlights
# ===========================================================================
slide4 = create_slide()
content_top = add_header(slide4, "What We Are Building This Semester", "Semester Roadmap")

roadmaps = [
    {
        "num": "PHASE 01",
        "title": "Hands-on Weekly Labs",
        "color": CYAN,
        "desc": "Guided interactive sessions covering Linux fundamentals, Burp Suite proxy intercept, packet captures, and exploit isolation in sandbox environments."
    },
    {
        "num": "PHASE 02",
        "title": "Campus Defense Initiative",
        "color": AMBER,
        "desc": "Campus-wide Security Open House featuring phishing awareness simulations, password auditing clinics, and live social engineering defense demonstrations."
    },
    {
        "num": "PHASE 03",
        "title": "BaySec CTF 2026",
        "color": GREEN_ACCENT,
        "desc": "Our flagship 24-hour campus Capture-The-Flag competition with dedicated beginner tracks, jeopardy challenges, scoreboards, and sponsor tech prizes."
    }
]

step_w = Inches(3.64)
step_h = Inches(4.7)

for i, rm in enumerate(roadmaps):
    card = slide4.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left_base + i * (step_w + gap),
        content_top + Inches(0.2),
        step_w,
        step_h
    )
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = rm["color"]
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.35)
    tf.margin_right = Inches(0.35)
    tf.margin_top = Inches(0.4)

    p = tf.paragraphs[0]
    p.text = rm["num"]
    p.font.name = FONT_CODE
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = rm["color"]

    p = tf.add_paragraph()
    p.text = rm["title"]
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.space_before = Pt(6)
    p.space_after = Pt(20)

    p = tf.add_paragraph()
    p.text = rm["desc"]
    p.font.name = FONT_BODY
    p.font.size = Pt(13.5)
    p.font.color.rgb = SLATE_LIGHT
    p.space_before = Pt(8)

add_speaker_notes(
    slide4,
    "Every session we host is designed to be hands-on. Instead of dry lectures, "
    "we run live walkthroughs, simulate defenses, and build practical skills. Later this term, "
    "we’ll host a campus-wide Security Open House and cap it off with our own BaySec CTF."
)

# ===========================================================================
# SLIDE 5: Free Practice Toolkit
# ===========================================================================
slide5 = create_slide()
content_top = add_header(slide5, "Your Free Security Sandbox", "Skill Acceleration")

# Sub-header note
sub = slide5.shapes.add_textbox(Inches(0.8), content_top - Inches(0.3), Inches(11.733), Inches(0.4))
p = sub.text_frame.paragraphs[0]
p.text = "No paid licenses or expensive subscriptions required to become formidable."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = BLUE_ACCENT

tool_w = Inches(3.64)
tool_h = Inches(3.7)

tools = [
    {
        "name": "TryHackMe (Free Tier)",
        "tag": "WEB & LABS",
        "color": RED_ACCENT,
        "desc": "Browser-accessible security rooms and guided capture-the-flag paths for solo practice. Perfect for learning web attack basics and networking fundamentals without configuring local VMs."
    },
    {
        "name": "OverTheWire (Bandit)",
        "tag": "LINUX MASTERY",
        "color": GREEN_ACCENT,
        "desc": "Gamified, interactive SSH challenge rooms to master Linux command-line essentials. Levels force you to dig through permissions, pipes, hex dumps, and shell redirection."
    },
    {
        "name": "CyberChef",
        "tag": "THE SWISS ARMY KNIFE",
        "color": CYAN,
        "desc": "The browser-based utility by GCHQ for decoding ciphers, Base64 strings, calculating hashes, and unpacking obfuscated malware payloads with simple drag-and-drop operations."
    }
]

for i, t in enumerate(tools):
    card = slide5.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left_base + i * (tool_w + gap),
        content_top + Inches(0.3),
        tool_w,
        tool_h
    )
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = t["color"]
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.35)

    p = tf.paragraphs[0]
    p.text = t["tag"]
    p.font.name = FONT_CODE
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = t["color"]

    p = tf.add_paragraph()
    p.text = t["name"]
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.space_before = Pt(4)
    p.space_after = Pt(14)

    p = tf.add_paragraph()
    p.text = t["desc"]
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = SLATE_LIGHT

# Bottom Callout Box
callout = slide5.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    Inches(6.05),
    Inches(11.733),
    Inches(0.8)
)
callout.fill.solid()
callout.fill.fore_color.rgb = RGBColor(15, 23, 42)
callout.line.color.rgb = CYAN
callout.line.width = Pt(1.5)
tf_c = callout.text_frame
tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf_c.paragraphs[0]
p.text = "💡 SFBU provides the lab space; these free platforms give you unlimited sandbox reps outside meetings."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide5,
    "Security is a rep-based craft. You don’t need expensive certifications or paid subscriptions "
    "to start. TryHackMe's free tier, OverTheWire Bandit, and CyberChef are all you need to build muscle memory right now."
)

# ===========================================================================
# SLIDE 6: Challenge 0x01 — The Leaked Credential
# ===========================================================================
slide6 = create_slide()
content_top = add_header(slide6, "Challenge 0x01: The Leaked Credential", "Live Hands-On CTF")

# Terminal Graphic Frame
term_box = slide6.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    content_top,
    Inches(8.2),
    Inches(5.2)
)
term_box.fill.solid()
term_box.fill.fore_color.rgb = RGBColor(15, 23, 42)
term_box.line.color.rgb = RED_ACCENT
term_box.line.width = Pt(2)

tf_t = term_box.text_frame
tf_t.word_wrap = True
tf_t.margin_left = Inches(0.45)
tf_t.margin_right = Inches(0.45)
tf_t.margin_top = Inches(0.35)

p = tf_t.paragraphs[0]
p.text = "[!] INCIDENT SCENARIO"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = RED_ACCENT

p = tf_t.add_paragraph()
p.text = "An unauthorized operator left their credentials cached in plain sight on our public test gateway."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(18)

p = tf_t.add_paragraph()
p.text = "[*] OBJECTIVE"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf_t.add_paragraph()
p.text = "Inspect the target surface, extract the hidden string, and decode the payload."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(18)

p = tf_t.add_paragraph()
p.text = "[#] FLAG FORMAT"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT

p = tf_t.add_paragraph()
p.text = "sfbu{...}"
p.font.name = FONT_CODE
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(20)

p = tf_t.add_paragraph()
p.text = "TARGET GATEWAY: https://MynameisKoi.github.io/baysec/week4/"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.color.rgb = BLUE_ACCENT

# Side Card: Live QR + Countdown Timer
side_card = slide6.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(9.3),
    content_top,
    Inches(3.233),
    Inches(5.2)
)
side_card.fill.solid()
side_card.fill.fore_color.rgb = CARD_BG
side_card.line.color.rgb = CYAN
side_card.line.width = Pt(1.5)

# Timer Badge on top of right side
timer_badge = slide6.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(9.65),
    content_top + Inches(0.3),
    Inches(2.533),
    Inches(0.9)
)
timer_badge.fill.solid()
timer_badge.fill.fore_color.rgb = RGBColor(30, 41, 59)
timer_badge.line.color.rgb = RED_ACCENT
timer_badge.line.width = Pt(2)
tf_time = timer_badge.text_frame
p = tf_time.paragraphs[0]
p.text = "TIME REMAINING"
p.font.name = FONT_CODE
p.font.size = Pt(10)
p.font.color.rgb = SLATE_MUTED
p.alignment = PP_ALIGN.CENTER
p = tf_time.add_paragraph()
p.text = "05:00"
p.font.name = FONT_CODE
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = RED_ACCENT
p.alignment = PP_ALIGN.CENTER

# QR Code in right card
slide6.shapes.add_picture(
    QR_CHALLENGE,
    Inches(9.8),
    content_top + Inches(1.5),
    Inches(2.233),
    Inches(2.233)
)

qr_txt = slide6.shapes.add_textbox(Inches(9.4), content_top + Inches(3.9), Inches(3.0), Inches(0.9))
tf_q = qr_txt.text_frame
tf_q.word_wrap = True
p = tf_q.paragraphs[0]
p.text = "Scan to Open Gateway"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = CYAN
p.alignment = PP_ALIGN.CENTER
p = tf_q.add_paragraph()
p.text = "or visit tinyurl.com/baysec-week4"
p.font.name = FONT_CODE
p.font.size = Pt(10)
p.font.color.rgb = SLATE_MUTED
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide6,
    "Let's put this into practice immediately. Pull out your laptop or phone and go to the URL on screen. "
    "An operator left an authorization token lingering inside this page. Your mission is to find it and extract the flag. "
    "You have 5 minutes—go!"
)

# ===========================================================================
# SLIDE 7: Challenge Solution Walkthrough
# ===========================================================================
slide7 = create_slide()
content_top = add_header(slide7, "Challenge 0x01: Rules & Submission", "League Competition")

col_w = Inches(5.6)
col_h = Inches(4.3)

# Left Column: Points & Submission Details
card_s1 = slide7.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    content_top,
    col_w,
    col_h
)
card_s1.fill.solid()
card_s1.fill.fore_color.rgb = CARD_BG
card_s1.line.color.rgb = CYAN
card_s1.line.width = Pt(1.5)

tf = card_s1.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.35)

p = tf.paragraphs[0]
p.text = "🚩 EARN LEAGUE POINTS"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf.add_paragraph()
p.text = "• Point Value: +100 Points on Master Leaderboard"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.space_before = Pt(12)

p = tf.add_paragraph()
p.text = "• Flag Format: sfbu{...}"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)

p = tf.add_paragraph()
p.text = "• Challenge Portal: tinyurl.com/baysec-week4"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)

p = tf.add_paragraph()
p.text = "• Submission Window: Open 24/7 until Week 5 Kickoff"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "💡 Hint: Inspect client-side source code carefully. Encoding is not encryption!"
p.font.name = FONT_BODY
p.font.size = Pt(12)
p.font.italic = True
p.font.color.rgb = SLATE_LIGHT

# Right Column: Solution Postponement & Leaderboard
card_s2 = slide7.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(6.933),
    content_top,
    col_w,
    col_h
)
card_s2.fill.solid()
card_s2.fill.fore_color.rgb = CARD_BG
card_s2.line.color.rgb = GREEN_ACCENT
card_s2.line.width = Pt(1.5)

tf = card_s2.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.35)

p = tf.paragraphs[0]
p.text = "⏳ SOLUTION REVEAL & DEBRIEF"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT

p = tf.add_paragraph()
p.text = "• Full Walkthrough Next Week:"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = AMBER
p.space_before = Pt(12)

p = tf.add_paragraph()
p.text = "The complete step-by-step exploit breakdown and source inspection solution will be revealed at the START OF NEXT WEEK'S SESSION (Week 5)!"
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "• Lock In Your Points:"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf.add_paragraph()
p.text = "Submit before next week to claim your points and advance your tier ranking."
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "TRACK LIVE STANDINGS: baysec/leaderboard.html"
p.font.name = FONT_CODE
p.font.size = Pt(11.5)
p.font.bold = True
p.font.color.rgb = CYAN

# Bottom Key Takeaway Callout
takeaway = slide7.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    Inches(6.25),
    Inches(11.733),
    Inches(0.65)
)
takeaway.fill.solid()
takeaway.fill.fore_color.rgb = RGBColor(15, 23, 42)
takeaway.line.color.rgb = CYAN
takeaway.line.width = Pt(1.5)
tf_take = takeaway.text_frame
tf_take.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf_take.paragraphs[0]
p.text = "🏆 Every challenge solved earns points toward end-of-semester Sentinel & Grandmaster prizes!"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide7,
    "If you solved Challenge 0x01, submit your token on the portal to secure your 100 points! "
    "If you haven't cracked it yet, you have all week until our next meeting. "
    "We will reveal and walk through the complete step-by-step solution at the start of Week 5."
)

# ===========================================================================
# SLIDE 8: Next Steps & Open Floor
# ===========================================================================
slide8 = create_slide()
content_top = add_header(slide8, "Get Connected & Next Steps", "Looking Ahead")

# Left Column: Next Meeting & Action Items
left_info = slide8.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.8),
    content_top,
    Inches(7.2),
    Inches(5.1)
)
left_info.fill.solid()
left_info.fill.fore_color.rgb = CARD_BG
left_info.line.color.rgb = CYAN
left_info.line.width = Pt(1.5)

tf = left_info.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.35)

p = tf.paragraphs[0]
p.text = "🚀 NEXT WEEK: LAB 01"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf.add_paragraph()
p.text = "Linux CLI Essentials & Wireshark Traffic Analysis"
p.font.name = FONT_HEADING
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "📅 MEETING SCHEDULE"
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = BLUE_ACCENT

p = tf.add_paragraph()
p.text = "Weekly 60-Minute Sessions (Final weekday/time confirmed via WhatsApp & Email)"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.color.rgb = SLATE_LIGHT
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "⚡ ACTION ITEM"
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf.add_paragraph()
p.text = "Make sure you are in the WhatsApp Community for room announcements!"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.color.rgb = WHITE
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "📫 CONTACTS"
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT

p = tf.add_paragraph()
p.text = "• Khoi Duong (President) — mduong@student.sfbu.edu"
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE

p = tf.add_paragraph()
p.text = "• Joed Prudente (Faculty Advisor) — joed.prudente@sfbu.edu"
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE

# Right Column: QR Code + Questions
right_qr = slide8.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(8.3),
    content_top,
    Inches(4.233),
    Inches(5.1)
)
right_qr.fill.solid()
right_qr.fill.fore_color.rgb = CARD_BG
right_qr.line.color.rgb = GREEN_ACCENT
right_qr.line.width = Pt(1.5)

tf = right_qr.text_frame
tf.word_wrap = True
tf.margin_top = Inches(0.35)
p = tf.paragraphs[0]
p.text = "JOIN THE COMMUNITY"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.alignment = PP_ALIGN.CENTER

slide8.shapes.add_picture(
    QR_WHATSAPP,
    Inches(9.266),
    content_top + Inches(0.8),
    Inches(2.3),
    Inches(2.3)
)

q_box = slide8.shapes.add_textbox(Inches(8.4), content_top + Inches(3.3), Inches(4.033), Inches(1.3))
tf = q_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "QUESTIONS?"
p.font.name = FONT_HEADING
p.font.size = Pt(28)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

p = tf.add_paragraph()
p.text = "The floor is now open for Q&A, setup help, or just chatting security."
p.font.name = FONT_BODY
p.font.size = Pt(12)
p.font.color.rgb = SLATE_LIGHT
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide8,
    "Next week, we kick off Lab 01 with Linux CLI survival skills and packet sniffing in Wireshark. "
    "Watch WhatsApp for the locked-in room and time slot. Thank you for coming out, and let’s open up the floor for any questions!"
)

# ---------------------------------------------------------------------------
# Save Presentation
# ---------------------------------------------------------------------------
try:
    prs.save(OUTPUT_PPTX)
    print(f"Presentation successfully saved to: {OUTPUT_PPTX}")
except PermissionError:
    fallback = os.path.join(SCRIPT_DIR, "BaySec_Week4_Kickoff_updated.pptx")
    prs.save(fallback)
    print(f"Notice: '{OUTPUT_PPTX}' is currently open in PowerPoint.")
    print(f"Saved updated presentation to: {fallback}")
