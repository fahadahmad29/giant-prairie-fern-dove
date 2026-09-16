#!/usr/bin/env python3
"""Professional academic figures for the MDPS project report."""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/workspace/artifacts/report_figures")
OUT.mkdir(parents=True, exist_ok=True)

NAVY = (27, 54, 93)
NAVY_D = (15, 32, 56)
TEAL = (15, 118, 110)
GOLD = (184, 148, 74)
SLATE = (51, 65, 85)
MUTED = (100, 116, 139)
LIGHT = (241, 245, 249)
WHITE = (255, 255, 255)
SOFT_TEAL = (204, 251, 241)
SOFT_BLUE = (219, 234, 254)
SOFT_GOLD = (254, 243, 199)
SOFT_RED = (254, 226, 226)
SOFT_GREEN = (220, 252, 231)
LINE = (203, 213, 225)
BLACK = (15, 23, 42)

SANS = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
SANS_B = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
SANS_I = "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf"


def font(path, size):
    return ImageFont.truetype(path, size)


def wrap(draw, text, fnt, max_w):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if draw.textlength(trial, font=fnt) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [text]


def rounded(draw, xy, r, fill, outline=None, width=2):
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def center_text(draw, xy, text, fnt, fill=WHITE):
    x1, y1, x2, y2 = xy
    lines = text.split("\n")
    lh = fnt.getbbox("Ag")[3] - fnt.getbbox("Ag")[1] + 4
    total = lh * len(lines)
    y = y1 + (y2 - y1 - total) / 2
    for line in lines:
        tw = draw.textlength(line, font=fnt)
        draw.text(((x1 + x2 - tw) / 2, y), line, font=fnt, fill=fill)
        y += lh


def text_box(draw, xy, title, subtitle=None, fill=NAVY, outline=NAVY, title_size=22, sub_size=16, radius=16, title_fill=WHITE, sub_fill=None):
    rounded(draw, xy, radius, fill=fill, outline=outline, width=3)
    tf = font(SANS_B, title_size)
    if subtitle:
        sf = font(SANS, sub_size)
        sub_fill = sub_fill or (255, 255, 255, )
        # slightly off-white subtitle
        lines_t = title.split("\n")
        lines_s = subtitle.split("\n")
        lh_t = tf.getbbox("Ag")[3] - tf.getbbox("Ag")[1] + 2
        lh_s = sf.getbbox("Ag")[3] - sf.getbbox("Ag")[1] + 2
        total = lh_t * len(lines_t) + 6 + lh_s * len(lines_s)
        x1, y1, x2, y2 = xy
        y = y1 + (y2 - y1 - total) / 2
        for line in lines_t:
            tw = draw.textlength(line, font=tf)
            draw.text(((x1 + x2 - tw) / 2, y), line, font=tf, fill=title_fill)
            y += lh_t
        y += 6
        for line in lines_s:
            tw = draw.textlength(line, font=sf)
            draw.text(((x1 + x2 - tw) / 2, y), line, font=sf, fill=(226, 232, 240) if fill == NAVY or fill == TEAL or fill == NAVY_D else SLATE)
            y += lh_s
    else:
        center_text(draw, xy, title, tf, fill=title_fill)


def arrow_right(draw, x1, y, x2, color=NAVY, w=4):
    draw.line((x1, y, x2 - 12, y), fill=color, width=w)
    draw.polygon([(x2, y), (x2 - 16, y - 8), (x2 - 16, y + 8)], fill=color)


def arrow_down(draw, x, y1, y2, color=NAVY, w=4):
    draw.line((x, y1, x, y2 - 12), fill=color, width=w)
    draw.polygon([(x, y2), (x - 8, y2 - 16), (x + 8, y2 - 16)], fill=color)


def caption_bar(draw, w, title):
    draw.rectangle((0, 0, w, 70), fill=NAVY)
    draw.rectangle((0, 70, w, 76), fill=GOLD)
    f = font(SANS_B, 28)
    tw = draw.textlength(title, font=f)
    draw.text(((w - tw) / 2, 18), title, font=f, fill=WHITE)


def save(img, name):
    path = OUT / name
    img.save(path, "PNG", optimize=True)
    print("wrote", path, img.size)
    return path


def fig_context():
    W, H = 1600, 780
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 1  |  Context Diagram (Level 0)")
    # actor
    text_box(d, (80, 300, 340, 480), "USER / HEALTH\nSEEKER", "Enters clinical or\nvoice-related values", fill=NAVY)
    arrow_right(d, 350, 390, 500)
    d.text((360, 350), "Input values", font=font(SANS_I, 16), fill=MUTED)
    text_box(d, (500, 250, 1100, 530), "MULTIPLE DISEASE\nPREDICTION APP", "Streamlit UI  •  Input conversion\nModel loading  •  Inference", fill=TEAL)
    arrow_right(d, 1110, 390, 1260)
    d.text((1115, 350), "Feature vector", font=font(SANS_I, 16), fill=MUTED)
    text_box(d, (1260, 250, 1520, 400), "DISEASE-SPECIFIC\nML MODEL", ".sav artifact", fill=NAVY)
    # return arrow
    d.line((1390, 410, 1390, 560, 800, 560, 800, 540), fill=GOLD, width=4)
    d.polygon([(800, 530), (792, 548), (808, 548)], fill=GOLD)
    d.text((980, 575), "Prediction result", font=font(SANS_I, 16), fill=MUTED)
    d.text((80, 700), "The same context-level flow applies to Diabetes, Heart Disease and Parkinson's; only the feature set and loaded model change.",
           font=font(SANS, 16), fill=SLATE)
    return save(img, "fig1_context.png")


def fig_existing_vs_proposed():
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 2  |  Existing Approach versus Proposed System")

    # left panel
    rounded(d, (50, 110, 770, 850), 20, fill=LIGHT, outline=LINE, width=2)
    d.text((90, 130), "TRADITIONAL / FRAGMENTED WORKFLOW", font=font(SANS_B, 20), fill=(185, 28, 28))
    text_box(d, (280, 190, 540, 280), "USER", fill=NAVY, title_size=22)
    # three tools
    boxes = [("DISEASE A TOOL", 140), ("DISEASE B TOOL", 390), ("DISEASE C TOOL", 640)]
    for title, x in boxes:
        text_box(d, (x, 400, x + 200, 520), title, fill=SOFT_RED, outline=(185, 28, 28), title_size=16, title_fill=NAVY_D)
        arrow_down(d, x + 100, 290, 390, color=(185, 28, 28))
    d.text((90, 580), "Different interfaces  •  Repeated navigation", font=font(SANS, 16), fill=SLATE)
    d.text((90, 610), "Inconsistent interaction patterns", font=font(SANS, 16), fill=SLATE)
    d.text((90, 660), "User must learn and operate a separate", font=font(SANS, 16), fill=MUTED)
    d.text((90, 690), "tool for each disease prediction task.", font=font(SANS, 16), fill=MUTED)

    # right panel
    rounded(d, (830, 110, 1550, 850), 20, fill=(240, 253, 250), outline=TEAL, width=2)
    d.text((870, 130), "PROPOSED UNIFIED SYSTEM", font=font(SANS_B, 20), fill=TEAL)
    text_box(d, (1060, 190, 1320, 270), "USER", fill=NAVY, title_size=22)
    arrow_down(d, 1190, 280, 340)
    text_box(d, (980, 345, 1400, 445), "STREAMLIT INTERFACE", "Sidebar navigation  •  One application", fill=TEAL, title_size=20)
    arrow_down(d, 1190, 455, 510)
    mods = [("DIABETES", 860), ("HEART", 1075), ("PARKINSON'S", 1290)]
    for title, x in mods:
        text_box(d, (x, 520, x + 200, 610), title, fill=NAVY, title_size=16)
    # merge arrows
    d.line((960, 620, 960, 670, 1190, 670, 1190, 690), fill=NAVY, width=3)
    d.line((1175, 620, 1175, 670), fill=NAVY, width=3)
    d.line((1390, 620, 1390, 670, 1190, 670), fill=NAVY, width=3)
    arrow_down(d, 1190, 670, 710)
    text_box(d, (1000, 715, 1380, 810), "PREDICTION RESULT", "Clear classification-style message", fill=GOLD, title_size=18, title_fill=NAVY_D, sub_fill=SLATE)
    return save(img, "fig2_existing_vs_proposed.png")


def fig_architecture():
    W, H = 1600, 980
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 3  |  Logical System Architecture")

    layers = [
        (110, "END USER", "Enters health parameters through the browser", NAVY),
        (230, "STREAMLIT UI & NAVIGATION", "option_menu sidebar  •  Diabetes / Heart / Parkinson's pages", TEAL),
        (350, "INPUT VALIDATION / CAST", "Ordered feature list  •  float() conversion of form values", (37, 99, 135)),
        (470, "MODEL ARTIFACTS", "diabetes_model.sav   •   heart_disease_model.sav   •   parkinsons_model.sav", NAVY_D),
        (590, "MODEL PREDICTION", "Loaded estimator.predict([user_input])", TEAL),
        (710, "RESULT DISPLAY", "st.success() classification message shown to the user", GOLD),
    ]
    for y, title, sub, fill in layers:
        text_box(d, (180, y, 1420, y + 95), title, sub, fill=fill, title_size=24, sub_size=16,
                 title_fill=NAVY_D if fill == GOLD else WHITE)
        if y != 710:
            arrow_down(d, 800, y + 97, y + 118, color=NAVY)

    d.text((180, 840), "Startup behaviour: pickle.load() reads the three saved model files from saved_models/", font=font(SANS, 18), fill=SLATE)
    d.text((180, 875), "before any user interaction. Each disease page then reuses the already-loaded estimator.", font=font(SANS, 18), fill=SLATE)
    d.text((180, 920), "Application language: Python    UI runtime: Streamlit    Persistence: pickle .sav artifacts", font=font(SANS_I, 16), fill=MUTED)
    return save(img, "fig3_architecture.png")


def fig_methodology():
    W, H = 1600, 620
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 4  |  End-to-End Inference Methodology")
    steps = [
        ("1", "Launch App", "Streamlit starts\nthe application UI"),
        ("2", "Select Module", "Sidebar chooses\none disease"),
        ("3", "Enter Inputs", "Disease-specific\nfeatures are filled"),
        ("4", "Convert Data", "Inputs are cast\nto float values"),
        ("5", "Run Model", "Loaded model\npredicts output"),
        ("6", "Show Result", "Diagnosis message\nis displayed"),
    ]
    x = 50
    for i, (num, title, sub) in enumerate(steps):
        # number badge
        d.ellipse((x + 85, 120, x + 145, 180), fill=GOLD, outline=NAVY, width=2)
        center_text(d, (x + 85, 120, x + 145, 180), num, font(SANS_B, 26), fill=NAVY_D)
        text_box(d, (x, 210, x + 230, 430), title, sub, fill=NAVY, title_size=20, sub_size=15)
        if i < len(steps) - 1:
            arrow_right(d, x + 238, 320, x + 258, color=GOLD, w=5)
        x += 258
    d.text((50, 500), "Control flow: the same high-level inference pattern is used for all three disease modules.", font=font(SANS, 18), fill=SLATE)
    d.text((50, 540), "Only the input schema and the loaded model artifact differ from one page to another.", font=font(SANS, 18), fill=SLATE)
    return save(img, "fig4_methodology.png")


def fig_ml_pipeline():
    W, H = 1600, 720
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 5  |  Machine-Learning Training and Inference Pipeline")

    # training row
    d.text((60, 110), "A. TRAINING (notebooks)", font=font(SANS_B, 20), fill=TEAL)
    tsteps = [
        ("Dataset CSV", "diabetes.csv\nheart.csv\nparkinsons.csv"),
        ("Split X / Y", "Features vs\ntarget label"),
        ("80 : 20 split", "stratify where\napplicable"),
        ("Fit estimator", "SVM / Logistic\nRegression"),
        ("Evaluate", "Train & test\naccuracy"),
        ("pickle.dump", ".sav model\nartifact"),
    ]
    x = 50
    for i, (t, s) in enumerate(tsteps):
        text_box(d, (x, 155, x + 220, 310), t, s, fill=TEAL, title_size=18, sub_size=14)
        if i < len(tsteps) - 1:
            arrow_right(d, x + 226, 232, x + 248)
        x += 258

    # inference row
    d.text((60, 360), "B. INFERENCE (app.py)", font=font(SANS_B, 20), fill=NAVY)
    isteps = [
        ("Raw inputs", "st.text_input()\nwidgets"),
        ("Ordered vector", "user_input list\nin feature order"),
        ("float() cast", "Numeric conversion\nbefore predict"),
        ("Loaded model", "pickle.load()\nat startup"),
        ("predict([x])", "Binary class\nlabel 0 / 1"),
        ("Message", "st.success()\ndiagnosis text"),
    ]
    x = 50
    for i, (t, s) in enumerate(isteps):
        text_box(d, (x, 405, x + 220, 560), t, s, fill=NAVY, title_size=18, sub_size=14)
        if i < len(isteps) - 1:
            arrow_right(d, x + 226, 482, x + 248)
        x += 258

    d.text((60, 610), "Training is performed once in Jupyter/Colab notebooks. The Streamlit app never retrains; it only loads saved estimators and runs predict().",
           font=font(SANS, 17), fill=SLATE)
    return save(img, "fig5_ml_pipeline.png")


def fig_flowchart():
    W, H = 1100, 1500
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 6  |  Application Flow Chart")

    def box(y1, y2, title, sub, fill=NAVY):
        text_box(d, (280, y1, 820, y2), title, sub, fill=fill, title_size=22, sub_size=16)

    def diamond(cy, text):
        # diamond for decision
        cx = 550
        pts = [(cx, cy - 70), (cx + 230, cy), (cx, cy + 70), (cx - 230, cy)]
        d.polygon(pts, fill=SOFT_GOLD, outline=GOLD, width=3)
        center_text(d, (cx - 180, cy - 40, cx + 180, cy + 40), text, font(SANS_B, 18), fill=NAVY_D)

    box(100, 190, "START", "streamlit run app.py", fill=TEAL)
    arrow_down(d, 550, 195, 230)
    box(230, 330, "Load three .sav models", "pickle.load at startup")
    arrow_down(d, 550, 335, 370)
    box(370, 470, "Render sidebar menu", "Diabetes / Heart / Parkinson's", fill=TEAL)
    arrow_down(d, 550, 475, 510)
    box(510, 610, "Open selected disease page", "Disease-specific input form")
    arrow_down(d, 550, 615, 650)
    box(650, 750, "User enters feature values", "st.text_input fields")
    arrow_down(d, 550, 755, 800)
    diamond(870, "Test Result\nbutton clicked?")
    d.line((320, 870, 140, 870), fill=MUTED, width=3)
    d.text((145, 835), "No — wait for input", font=font(SANS_B, 15), fill=MUTED)
    d.line((140, 870, 140, 700, 270, 700), fill=MUTED, width=3)
    d.polygon([(270, 700), (254, 692), (254, 708)], fill=MUTED)
    arrow_down(d, 550, 940, 1015, color=NAVY)
    d.text((580, 945), "Yes", font=font(SANS_B, 16), fill=NAVY)
    box(1015, 1120, "Cast inputs to float", "Build ordered user_input list", fill=(37, 99, 135))
    arrow_down(d, 550, 1125, 1160)
    box(1160, 1260, "model.predict([user_input])", "Class label 0 or 1")
    arrow_down(d, 550, 1265, 1300)
    box(1300, 1400, "Display diagnosis message", "st.success(...)  •  END", fill=GOLD)
    return save(img, "fig6_flowchart.png")


def fig_usecase():
    W, H = 1500, 980
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 7  |  Use Case Diagram")

    # actor
    # stick figure
    ax, ay = 180, 430
    d.ellipse((ax - 28, ay - 170, ax + 28, ay - 114), outline=NAVY, width=4)
    d.line((ax, ay - 114, ax, ay - 20), fill=NAVY, width=4)
    d.line((ax - 50, ay - 80, ax + 50, ay - 80), fill=NAVY, width=4)
    d.line((ax, ay - 20, ax - 40, ay + 50), fill=NAVY, width=4)
    d.line((ax, ay - 20, ax + 40, ay + 50), fill=NAVY, width=4)
    d.text((ax - 90, ay + 70), "Health User /\nStudent Tester", font=font(SANS_B, 18), fill=NAVY)

    # system boundary
    rounded(d, (380, 120, 1420, 900), 18, fill=WHITE, outline=NAVY, width=3)
    d.text((700, 140), "Multiple Disease Prediction System", font=font(SANS_B, 22), fill=NAVY)

    cases = [
        (520, 230, 900, 330, "Select disease module"),
        (520, 360, 900, 460, "Enter health / voice inputs"),
        (520, 490, 900, 590, "Request diabetes prediction"),
        (520, 620, 900, 720, "Request heart prediction"),
        (520, 750, 900, 850, "Request Parkinson's prediction"),
        (980, 360, 1360, 460, "Load saved ML models"),
        (980, 520, 1360, 620, "Run model.predict()"),
        (980, 680, 1360, 780, "View diagnosis message"),
    ]
    for x1, y1, x2, y2, t in cases:
        # ellipse-like use case
        d.ellipse((x1, y1, x2, y2), fill=SOFT_BLUE, outline=NAVY, width=2)
        center_text(d, (x1, y1, x2, y2), t, font(SANS_B, 16), fill=NAVY_D)

    # actor to left use cases
    for y in (280, 410, 540, 670, 800):
        d.line((240, 400, 520, y), fill=SLATE, width=2)

    d.line((900, 410, 980, 410), fill=SLATE, width=2)
    d.line((900, 540, 980, 570), fill=SLATE, width=2)
    d.line((900, 670, 980, 570), fill=SLATE, width=2)
    d.line((900, 800, 980, 730), fill=SLATE, width=2)
    d.line((1170, 460, 1170, 520), fill=SLATE, width=2)
    d.line((1170, 620, 1170, 680), fill=SLATE, width=2)
    return save(img, "fig7_usecase.png")


def fig_hld_lld():
    W, H = 1600, 820
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 8  |  High-Level Design and Low-Level Design")

    d.text((80, 110), "HIGH-LEVEL DESIGN (HLD)", font=font(SANS_B, 22), fill=NAVY)
    hlds = [("UI /\nNAVIGATION", 70), ("DISEASE\nMODULES", 430), ("MODEL\nLAYER", 790), ("RESULT\nLAYER", 1150)]
    for title, x in hlds:
        text_box(d, (x, 160, x + 320, 320), title, fill=NAVY, title_size=22)
    for x in (390, 750, 1110):
        arrow_right(d, x, 240, x + 40, color=GOLD, w=5)

    d.text((80, 380), "LOW-LEVEL DESIGN (LLD)  —  one repeating implementation pattern", font=font(SANS_B, 22), fill=TEAL)
    llds = [
        ("Input widgets", "st.text_input()"),
        ("Feature list", "ordered user_input"),
        ("Conversion", "float(x) for x in list"),
        ("Inference", "model.predict([x])"),
        ("Output", "st.success(message)"),
    ]
    x = 50
    for i, (t, s) in enumerate(llds):
        text_box(d, (x, 440, x + 270, 620), t, s, fill=TEAL, title_size=20, sub_size=16)
        if i < len(llds) - 1:
            arrow_right(d, x + 278, 530, x + 308, color=GOLD)
        x += 318
    d.text((80, 680), "HLD describes the four logical layers of the deployed application. LLD maps each layer to the exact Streamlit / scikit-learn calls", font=font(SANS, 17), fill=SLATE)
    d.text((80, 715), "used in app.py. The same LLD sequence is reused independently for Diabetes (8 features), Heart Disease (13 features) and Parkinson's (22 features).", font=font(SANS, 17), fill=SLATE)
    return save(img, "fig8_hld_lld.png")


def fig_ui():
    W, H = 1500, 780
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    caption_bar(d, W, "Figure 9  |  User Interface and Navigation Design")

    # sidebar
    rounded(d, (60, 120, 430, 720), 16, fill=NAVY)
    d.text((90, 150), "SIDEBAR", font=font(SANS_B, 16), fill=GOLD)
    d.text((90, 185), "Multiple Disease\nPrediction System", font=font(SANS_B, 22), fill=WHITE)
    items = [
        (280, "Diabetes Prediction", True),
        (370, "Heart Disease Prediction", False),
        (460, "Parkinsons Prediction", False),
    ]
    for y, label, active in items:
        fill = TEAL if active else (36, 68, 110)
        rounded(d, (90, y, 400, y + 70), 12, fill=fill)
        d.text((110, y + 22), label, font=font(SANS_B, 16), fill=WHITE)

    # main
    rounded(d, (470, 120, 1440, 720), 16, fill=LIGHT, outline=LINE, width=2)
    d.text((510, 150), "DISEASE PAGE", font=font(SANS_B, 16), fill=TEAL)
    d.text((510, 185), "Diabetes Prediction using ML", font=font(SANS_B, 28), fill=NAVY)

    # form grid
    labels = ["Pregnancies", "Glucose", "Blood Pressure", "Skin Thickness", "Insulin", "BMI", "DPF", "Age"]
    for i, lab in enumerate(labels):
        r, c = divmod(i, 3)
        x = 510 + c * 290
        y = 250 + r * 90
        rounded(d, (x, y, x + 270, y + 70), 10, fill=WHITE, outline=LINE, width=2)
        d.text((x + 14, y + 10), lab, font=font(SANS, 13), fill=MUTED)
        d.text((x + 14, y + 34), "text input", font=font(SANS_I, 14), fill=SLATE)

    rounded(d, (510, 530, 900, 610), 12, fill=TEAL)
    center_text(d, (510, 530, 900, 610), "Diabetes Test Result", font(SANS_B, 20), WHITE)
    rounded(d, (510, 630, 1400, 700), 12, fill=SOFT_GREEN, outline=(21, 128, 61), width=2)
    center_text(d, (510, 630, 1400, 700), "The person is diabetic   /   The person is not diabetic", font(SANS_B, 18), (21, 128, 61))
    return save(img, "fig9_ui.png")


if __name__ == "__main__":
    fig_context()
    fig_existing_vs_proposed()
    fig_architecture()
    fig_methodology()
    fig_ml_pipeline()
    fig_flowchart()
    fig_usecase()
    fig_hld_lld()
    fig_ui()
    print("all figures done")
