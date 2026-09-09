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

OUTPUT_PPTX = os.path.join(SCRIPT_DIR, "BaySec_Week5_Foundations.pptx")

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
QR_CHALLENGE = make_qr("https://MynameisKoi.github.io/baysec/week5/", "qr_challenge.png")

BLUE_LOGO = os.path.join(ASSETS_DIR, "baysec-blue.png")
RED_LOGO = os.path.join(ASSETS_DIR, "baysec-red.png")

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT
blank_layout = prs.slide_layouts[6]

def create_slide():
    slide = prs.slides.add_slide(blank_layout)
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
            Inches(2.8),
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

tag = slide1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(4.666),
    Inches(0.55),
    Inches(4.0),
    Inches(0.38)
)
tag.fill.solid()
tag.fill.fore_color.rgb = RGBColor(15, 23, 42)
tag.line.color.rgb = CYAN
tag.line.width = Pt(1.5)
tf = tag.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "WEEK 5 | NETWORKING & TRAFFIC ANALYSIS"
p.font.name = FONT_CODE
p.font.size = Pt(10.5)
p.font.bold = True
p.font.color.rgb = CYAN
p.alignment = PP_ALIGN.CENTER

title_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.15), Inches(11.333), Inches(1.3))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Foundations: Linux CLI & Wireshark"
p.font.name = FONT_HEADING
p.font.size = Pt(40)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

sub_box = slide1.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.333), Inches(0.6))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.text = "BaySec @ SFBU • Packet Sniffing & Network Forensics Lab"
p.font.name = FONT_BODY
p.font.size = Pt(19)
p.font.color.rgb = BLUE_ACCENT
p.alignment = PP_ALIGN.CENTER

if os.path.exists(BLUE_LOGO) and os.path.exists(RED_LOGO):
    logo_w = Inches(2.2)
    logo_h = Inches(2.2)
    center_x = SLIDE_WIDTH / 2
    slide1.shapes.add_picture(BLUE_LOGO, center_x - Inches(2.35), Inches(3.0), logo_w, logo_h)
    slide1.shapes.add_picture(RED_LOGO, center_x + Inches(0.15), Inches(3.0), logo_w, logo_h)

wel_box = slide1.shapes.add_textbox(Inches(2.5), Inches(5.35), Inches(8.333), Inches(0.8))
tf = wel_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Welcome back! Please scan the QR code to log today's attendance (+50 League Points) while we get started."
p.font.name = FONT_BODY
p.font.size = Pt(14)
p.font.color.rgb = SLATE_LIGHT
p.alignment = PP_ALIGN.CENTER

# QR 1: Attendance
card_qr1 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.9), Inches(2.3), Inches(2.2))
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

# QR 2: WhatsApp
card_qr2 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.233), Inches(4.9), Inches(2.3), Inches(2.2))
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
    "Welcome to Week 5 of BaySec! Make sure to scan the QR code for today's attendance so your points get added to the league leaderboard. "
    "Today, we move from the kickoff into core foundations: mastering the Linux command line and learning how to inspect live packets using Wireshark."
)

# ===========================================================================
# SLIDE 2: Week 4 Recap & Standings
# ===========================================================================
slide2 = create_slide()
content_top = add_header(slide2, "Week 4 Recap & League Standings", "Season Progress")

col_w = Inches(5.6)
col_h = Inches(5.1)

card1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), content_top, col_w, col_h)
card1.fill.solid()
card1.fill.fore_color.rgb = CARD_BG
card1.line.color.rgb = CYAN
card1.line.width = Pt(1.5)
tf = card1.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.4)

p = tf.paragraphs[0]
p.text = "🚩 CHALLENGE 0x01 DEBRIEF"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf.add_paragraph()
p.text = "In our kickoff lab, we uncovered an authorization token exposed inside HTML comments on the client gateway:"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.color.rgb = WHITE
p.space_before = Pt(8)
p.space_after = Pt(10)

p = tf.add_paragraph()
p.text = "echo \"c2ZidXt3M2xjMG0z...\" | base64 -d"
p.font.name = FONT_CODE
p.font.size = Pt(11)
p.font.color.rgb = GREEN_ACCENT
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "• Solvers earned +100 Challenge Points\n• Total Kickoff potential: 150 pts (Attendance + Challenge)\n• Key Lesson: The browser is untrusted territory."
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = SLATE_LIGHT

# Right Card: Leaderboard Launch
card2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.933), content_top, col_w, col_h)
card2.fill.solid()
card2.fill.fore_color.rgb = CARD_BG
card2.line.color.rgb = AMBER
card2.line.width = Pt(1.5)
tf = card2.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.4)

p = tf.paragraphs[0]
p.text = "🏆 LIVE LEAGUE STANDINGS"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf.add_paragraph()
p.text = "Our official BaySec Leaderboard is now live on the portal!"
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.color.rgb = WHITE
p.space_before = Pt(8)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "CURRENT TIER LADDER:"
p.font.name = FONT_CODE
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = SLATE_MUTED

tiers = [
    "🌱 Initiate: 0 – 149 pts",
    "⚡ Operator: 150 – 299 pts (Current Leaders!)",
    "🛡️ Sentinel: 300 – 599 pts",
    "⚔️ Breaker: 600 – 999 pts",
    "👑 Grandmaster: 1,000+ pts"
]
for t in tiers:
    p = tf.add_paragraph()
    p.text = f"• {t}"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE

p = tf.add_paragraph()
p.text = "Check your standing anytime: baysec/leaderboard.html"
p.font.name = FONT_CODE
p.font.size = Pt(11)
p.font.color.rgb = CYAN
p.space_before = Pt(14)

add_speaker_notes(
    slide2,
    "Awesome job to everyone who solved Challenge 0x01 last week! You successfully broke into the Operator tier with 150 points. "
    "Our official live leaderboard is now up on the website. Today's Challenge 0x02 offers another 100 points, so there is plenty of room to climb the standings."
)

# ===========================================================================
# SLIDE 3: Linux CLI Survival Guide
# ===========================================================================
slide3 = create_slide()
content_top = add_header(slide3, "Linux Command-Line Survival Kit", "Terminal Mastery")

card_w = Inches(3.64)
card_h = Inches(4.3)
gap = Inches(0.4)
left_base = Inches(0.8)

tools_sections = [
    {
        "title": "FILESYSTEM & PERMS",
        "color": CYAN,
        "items": [
            ("ls -la", "List all files including hidden dotfiles & permissions"),
            ("cat / head / tail", "View start, end, or full text streams"),
            ("chmod +x", "Grant execute rights on scripts & binaries"),
            ("find / -name", "Locate files matching specific patterns")
        ]
    },
    {
        "title": "PIPES & WRANGLING",
        "color": GREEN_ACCENT,
        "items": [
            ("grep -rn \"flag\"", "Search recursively for patterns in files"),
            ("awk '{print $1}'", "Extract specific column outputs"),
            ("cut -d',' -f2", "Split delimited streams into fields"),
            ("sort | uniq -c", "Count occurrences of duplicate entries")
        ]
    },
    {
        "title": "NETWORK & SOCKETS",
        "color": AMBER,
        "items": [
            ("ip a / ifconfig", "Inspect interface IP & MAC addresses"),
            ("curl -v http://...", "Send raw HTTP requests & inspect headers"),
            ("ss -tulpn", "List listening TCP/UDP ports & socket PIDs"),
            ("tcpdump -i any", "Capture raw packets on terminal console")
        ]
    }
]

for i, sec in enumerate(tools_sections):
    card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_base + i * (card_w + gap), content_top, card_w, card_h)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = sec["color"]
    card.line.width = Pt(1.5)
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.35)

    p = tf.paragraphs[0]
    p.text = sec["title"]
    p.font.name = FONT_CODE
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = sec["color"]
    p.space_after = Pt(14)

    for cmd, desc in sec["items"]:
        p = tf.add_paragraph()
        p.text = f"$ {cmd}"
        p.font.name = FONT_CODE
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE
        
        p = tf.add_paragraph()
        p.text = desc
        p.font.name = FONT_BODY
        p.font.size = Pt(11)
        p.font.color.rgb = SLATE_MUTED
        p.space_after = Pt(8)

banner = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.65))
banner.fill.solid()
banner.fill.fore_color.rgb = RGBColor(15, 23, 42)
banner.line.color.rgb = BLUE_ACCENT
banner.line.width = Pt(1)
tf = banner.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "⚡ In security operations, the Linux shell is your primary weapon for triage, forensics, and exploit delivery."
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide3,
    "Security pros spend 90% of their time on the Linux command line. Piping commands together—like taking output from curl, "
    "passing it through grep, and filtering with awk—lets you dissect gigabytes of server logs or packet captures in seconds."
)

# ===========================================================================
# SLIDE 4: Networking Under the Hood
# ===========================================================================
slide4 = create_slide()
content_top = add_header(slide4, "Packets, Ports & Protocols", "Network Architecture")

col_w = Inches(5.6)
col_h = Inches(5.1)

card1 = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), content_top, col_w, col_h)
card1.fill.solid()
card1.fill.fore_color.rgb = CARD_BG
card1.line.color.rgb = CYAN
card1.line.width = Pt(1.5)
tf = card1.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.4)

p = tf.paragraphs[0]
p.text = "🌐 THE 4-LAYER PACKET ANATOMY"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

layers = [
    ("Layer 7 (Application)", "HTTP, DNS, SSH, TLS — The user data payload."),
    ("Layer 4 (Transport)", "TCP / UDP — Ports (e.g. 80, 443), sequence numbers, reliability."),
    ("Layer 3 (Network)", "IPv4 / IPv6 — Logical routing from Source IP to Destination IP."),
    ("Layer 2 (Data Link)", "Ethernet — MAC addresses for hop-by-hop local hardware framing.")
]
for title, desc in layers:
    p = tf.add_paragraph()
    p.text = title
    p.font.name = FONT_CODE
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = BLUE_ACCENT
    p.space_before = Pt(8)

    p = tf.add_paragraph()
    p.text = desc
    p.font.name = FONT_BODY
    p.font.size = Pt(11.5)
    p.font.color.rgb = SLATE_LIGHT

# Right Card: TCP 3-Way Handshake
card2 = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.933), content_top, col_w, col_h)
card2.fill.solid()
card2.fill.fore_color.rgb = CARD_BG
card2.line.color.rgb = RED_ACCENT
card2.line.width = Pt(1.5)
tf = card2.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.4)
tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.4)

p = tf.paragraphs[0]
p.text = "🤝 THE TCP 3-WAY HANDSHAKE"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = RED_ACCENT

p = tf.add_paragraph()
p.text = "Before any web request can fly, TCP establishes state:"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)
p.space_after = Pt(14)

steps = [
    ("1. SYN (Synchronize)", "Client ➔ Server: 'Hey, let's establish a session on port 80 (Seq=0)'"),
    ("2. SYN-ACK", "Server ➔ Client: 'Acknowledged, let's talk (Seq=0, Ack=1)'"),
    ("3. ACK", "Client ➔ Server: 'Connection confirmed (Ack=1)'"),
    ("4. DATA STREAM", "Payload flows (e.g. GET / POST HTTP requests)")
]
for title, desc in steps:
    p = tf.add_paragraph()
    p.text = title
    p.font.name = FONT_CODE
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = GREEN_ACCENT

    p = tf.add_paragraph()
    p.text = desc
    p.font.name = FONT_BODY
    p.font.size = Pt(11.5)
    p.font.color.rgb = SLATE_LIGHT
    p.space_after = Pt(8)

add_speaker_notes(
    slide4,
    "Packets are like Russian nesting dolls. An Ethernet frame carries an IP packet, which carries a TCP segment, which carries HTTP data. "
    "Every reliable connection on the internet starts with that 3-way handshake: SYN, SYN-ACK, ACK. When we look at Wireshark, we can see every step."
)

# ===========================================================================
# SLIDE 5: Wireshark Field Manual
# ===========================================================================
slide5 = create_slide()
content_top = add_header(slide5, "Wireshark Packet Analysis Manual", "Defensive Triage")

card_w = Inches(3.64)
card_h = Inches(4.3)

sections = [
    {
        "title": "DISPLAY FILTERS",
        "tag": "FILTERING NOISE",
        "color": GREEN_ACCENT,
        "items": [
            ("http", "Isolates plaintext HTTP traffic"),
            ("tcp.port == 8080", "Shows traffic on specific port"),
            ("ip.src == 192.168.1.42", "Tracks host requests"),
            ("dns || arp", "Inspects resolution handshakes")
        ]
    },
    {
        "title": "FOLLOW TCP STREAM",
        "tag": "REASSEMBLY",
        "color": CYAN,
        "items": [
            ("Right-Click Packet", "Select Follow ➔ TCP Stream"),
            ("Conversational View", "Reassembles fragmented chunks into human-readable text"),
            ("Color Code", "Red text = Client, Blue text = Server response"),
            ("Secret Revealer", "Extracts credentials & files directly")
        ]
    },
    {
        "title": "HTTP VS HTTPS",
        "tag": "THE PLAINTEXT RISK",
        "color": RED_ACCENT,
        "items": [
            ("HTTP (Port 80/8080)", "Plaintext: Headers, passwords, cookies visible to any observer"),
            ("HTTPS / TLS (Port 443)", "Encrypted: Payload appears as random pseudorandom bytes"),
            ("Packet Sniffing Risk", "On shared Wi-Fi/LAN, plaintext credentials are instant leaks")
        ]
    }
]

for i, sec in enumerate(sections):
    card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_base + i * (card_w + gap), content_top, card_w, card_h)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = sec["color"]
    card.line.width = Pt(1.5)
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.35)

    p = tf.paragraphs[0]
    p.text = sec["tag"]
    p.font.name = FONT_CODE
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = sec["color"]

    p = tf.add_paragraph()
    p.text = sec["title"]
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.space_after = Pt(14)

    for cmd, desc in sec["items"]:
        p = tf.add_paragraph()
        p.text = f"▸ {cmd}"
        p.font.name = FONT_CODE
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.name = FONT_BODY
        p.font.size = Pt(11)
        p.font.color.rgb = SLATE_MUTED
        p.space_after = Pt(6)

banner = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.65))
banner.fill.solid()
banner.fill.fore_color.rgb = RGBColor(15, 23, 42)
banner.line.color.rgb = CYAN
banner.line.width = Pt(1)
tf = banner.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "💡 Wireshark captures millions of packets. The skill of a SOC analyst is using display filters to eliminate noise."
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide5,
    "The most powerful feature for beginners in Wireshark is 'Follow TCP Stream.' "
    "Instead of looking at individual TCP fragments, Wireshark stitches the entire HTTP conversation back together into plain text. "
    "Let's put this into practice right now."
)

# ===========================================================================
# SLIDE 6: Challenge 0x02 — The Intercepted Uplink
# ===========================================================================
slide6 = create_slide()
content_top = add_header(slide6, "Challenge 0x02: The Intercepted Uplink", "Live Hands-On Lab")

term_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), content_top, Inches(8.2), Inches(5.2))
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
p.text = "A network tap on our internal test subnet intercepted an unencrypted communication stream. An operator transmitted an authentication beacon. Inspect the packet stream, isolate the protocol layers, and retrieve the authorization flag."
p.font.name = FONT_BODY
p.font.size = Pt(13.5)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf_t.add_paragraph()
p.text = "[*] OBJECTIVE & METHODS"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf_t.add_paragraph()
p.text = "• Method A: Use our web-based CyberShark Packet Analyzer on the portal.\n• Method B: Download traffic_analysis_0x02.pcap and inspect in Wireshark."
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf_t.add_paragraph()
p.text = "[#] FLAG FORMAT: sfbu{...}"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.space_after = Pt(16)

p = tf_t.add_paragraph()
p.text = "TARGET PORTAL: tinyurl.com/baysec-week5"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.color.rgb = BLUE_ACCENT

# Right Card: Timer + QR
side_card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.3), content_top, Inches(3.233), Inches(5.2))
side_card.fill.solid()
side_card.fill.fore_color.rgb = CARD_BG
side_card.line.color.rgb = CYAN
side_card.line.width = Pt(1.5)

timer_badge = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), content_top + Inches(0.3), Inches(2.533), Inches(0.9))
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

slide6.shapes.add_picture(QR_CHALLENGE, Inches(9.8), content_top + Inches(1.5), Inches(2.233), Inches(2.233))

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
p.text = "or visit tinyurl.com/baysec-week5"
p.font.name = FONT_CODE
p.font.size = Pt(10)
p.font.color.rgb = SLATE_MUTED
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide6,
    "Time for Challenge 0x02! Pull out your laptops or phones and scan the QR code. "
    "You can analyze the traffic directly in your browser using our CyberShark tool, or download the .pcap to open it in Wireshark. "
    "Find the operator's authentication token and submit the flag. 5 minutes on the clock—go!"
)

# ===========================================================================
# SLIDE 7: Challenge Solution Walkthrough
# ===========================================================================
slide7 = create_slide()
content_top = add_header(slide7, "Challenge 0x02: Rules & Submission", "League Competition")

col_w = Inches(5.6)
col_h = Inches(4.3)

card_s1 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), content_top, col_w, col_h)
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
p.text = "• Point Value: +150 Points on Master Leaderboard"
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
p.text = "• Target Portal: tinyurl.com/baysec-week5"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)

p = tf.add_paragraph()
p.text = "• Analysis Tools: In-browser CyberShark or raw .pcap download"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = WHITE
p.space_before = Pt(8)
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "💡 Hint: Filter for http or port 8080. Inspect the HTTP POST payload!"
p.font.name = FONT_BODY
p.font.size = Pt(12)
p.font.italic = True
p.font.color.rgb = SLATE_LIGHT

card_s2 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.933), content_top, col_w, col_h)
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
p.text = "The complete step-by-step packet dissection and forensic breakdown will be revealed at the START OF NEXT WEEK'S SESSION (Week 6)!"
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "• Advance Your Rank:"
p.font.name = FONT_CODE
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = AMBER

p = tf.add_paragraph()
p.text = "Submit before next week to claim your points and unlock the Operator or Sentinel tier badge."
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

takeaway = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.65))
takeaway.fill.solid()
takeaway.fill.fore_color.rgb = RGBColor(15, 23, 42)
takeaway.line.color.rgb = CYAN
takeaway.line.width = Pt(1.5)
tf_take = takeaway.text_frame
tf_take.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf_take.paragraphs[0]
p.text = "🏆 Every flag submitted pushes you closer to end-of-semester Sentinel & Grandmaster hardware prizes!"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide7,
    "If you found the flag in the uplink, submit it on the portal to lock in your +150 points! "
    "If you are still working through the capture, don't worry—the challenge is open all week. "
    "You can analyze it in CyberShark or download the PCAP for Wireshark. "
    "The full step-by-step solution and forensic breakdown will be revealed at the start of Week 6."
)

# ===========================================================================
# SLIDE 8: Next Steps & Open Floor
# ===========================================================================
slide8 = create_slide()
content_top = add_header(slide8, "Next Steps & Looking Ahead", "Looking Ahead")

left_info = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), content_top, Inches(7.2), Inches(5.1))
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
p.text = "🚀 NEXT WEEK: RECON & OSINT"
p.font.name = FONT_CODE
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf.add_paragraph()
p.text = "Active/Passive Reconnaissance, Nmap Network Scanning & Google Dorking"
p.font.name = FONT_HEADING
p.font.size = Pt(16.5)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_before = Pt(4)
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "📅 LEADERBOARD UPDATE"
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = BLUE_ACCENT

p = tf.add_paragraph()
p.text = "Points from today's attendance (+50) and Challenge 0x02 (+100) are being tallied. Check your tier progress on the leaderboard!"
p.font.name = FONT_BODY
p.font.size = Pt(13)
p.font.color.rgb = SLATE_LIGHT
p.space_after = Pt(14)

p = tf.add_paragraph()
p.text = "📫 CONTACTS"
p.font.name = FONT_CODE
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT

p = tf.add_paragraph()
p.text = "• Khoi Duong (President) — mduong@student.sfbu.edu\n• Joed Prudente (Faculty Advisor) — joed.prudente@sfbu.edu"
p.font.name = FONT_BODY
p.font.size = Pt(12.5)
p.font.color.rgb = WHITE

right_qr = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), content_top, Inches(4.233), Inches(5.1))
right_qr.fill.solid()
right_qr.fill.fore_color.rgb = CARD_BG
right_qr.line.color.rgb = GREEN_ACCENT
right_qr.line.width = Pt(1.5)
tf = right_qr.text_frame
tf.word_wrap = True
tf.margin_top = Inches(0.35)
p = tf.paragraphs[0]
p.text = "WHATSAPP COMMUNITY"
p.font.name = FONT_CODE
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = GREEN_ACCENT
p.alignment = PP_ALIGN.CENTER

slide8.shapes.add_picture(QR_WHATSAPP, Inches(9.266), content_top + Inches(0.8), Inches(2.3), Inches(2.3))

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
p.text = "The floor is now open for Q&A, lab troubleshooting, or discussing CTF strategies."
p.font.name = FONT_BODY
p.font.size = Pt(12)
p.font.color.rgb = SLATE_LIGHT
p.alignment = PP_ALIGN.CENTER

add_speaker_notes(
    slide8,
    "Next week we dive into Reconnaissance: scanning subnets with Nmap and hunting open-source intelligence with Google Dorks. "
    "Make sure you're in WhatsApp for lab files. Thank you everyone, let's open the floor for questions!"
)

prs.save(OUTPUT_PPTX)
print(f"Week 5 Presentation successfully saved to: {OUTPUT_PPTX}")
