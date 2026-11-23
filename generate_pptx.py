import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_human_grade_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 Widescreen
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Clean, High-Contrast Executive Dark Palette
    BG_DARK = RGBColor(11, 15, 25)        # #0B0F19 Dark Navy Slate
    CARD_BG = RGBColor(17, 24, 39)        # #111827 Dark Card Slate
    CARD_BORDER = RGBColor(31, 41, 55)    # #1F2937 Border
    ACCENT_BLUE = RGBColor(56, 189, 248)  # #38BDF8 Cyan
    ACCENT_ORANGE = RGBColor(249, 115, 22)# #F97316 Orange
    ACCENT_RED = RGBColor(239, 68, 68)    # #EF4444 Red
    ACCENT_GREEN = RGBColor(16, 185, 129) # #10B981 Emerald
    ACCENT_PURPLE = RGBColor(168, 85, 247)# #A855F7 Purple
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(156, 163, 175)   # Slate 400
    LIGHT_GRAY = RGBColor(226, 232, 240)  # Slate 200

    assets_dir = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\assets"

    def set_slide_background(slide, color=BG_DARK):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, title, presenter=""):
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(9.0), Inches(0.8))
        p_t = tb_title.text_frame.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.size = Pt(23)
        r_t.font.bold = True
        r_t.font.color.rgb = WHITE

        if presenter:
            tb_p = slide.shapes.add_textbox(Inches(9.2), Inches(0.48), Inches(3.3), Inches(0.6))
            p_p = tb_p.text_frame.paragraphs[0]
            p_p.alignment = PP_ALIGN.RIGHT
            r_p = p_p.add_run()
            r_p.text = f"🎙️ {presenter}"
            r_p.font.size = Pt(12)
            r_p.font.bold = True
            r_p.font.color.rgb = ACCENT_BLUE

    def create_card(slide, left, top, width, height, bg=CARD_BG, border=CARD_BORDER, lw=1.2):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = border
        card.line.width = Pt(lw)
        return card

    # =========================================================================
    # SLIDE 1: BIG IMPACT OPENING (Hero, High Contrast, Human Student Feel)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, RGBColor(7, 10, 19))

    # Main Hero Card
    hero = create_card(s1, Inches(0.8), Inches(0.8), Inches(11.733), Inches(4.3),
                       bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72), lw=2)

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.15), Inches(11.0), Inches(3.6))
    tf1 = tb1.text_frame

    p_badge = tf1.paragraphs[0]
    r_b = p_badge.add_run()
    r_b.text = "OPERATING SYSTEMS PBL SHOWCASE"
    r_b.font.size = Pt(12)
    r_b.font.bold = True
    r_b.font.color.rgb = ACCENT_ORANGE
    p_badge.space_after = Pt(10)

    p_main = tf1.add_paragraph()
    r_m = p_main.add_run()
    r_m.text = "CPU Scheduling Beyond Textbooks"
    r_m.font.size = Pt(36)
    r_m.font.bold = True
    r_m.font.color.rgb = WHITE
    p_main.space_after = Pt(6)

    p_sub = tf1.add_paragraph()
    r_s = p_sub.add_run()
    r_s.text = "Modeling Real Kernel Mechanics"
    r_s.font.size = Pt(20)
    r_s.font.bold = True
    r_s.font.color.rgb = ACCENT_BLUE
    p_sub.space_after = Pt(24)

    # 3 High-Impact Visual Badges
    p_pills = tf1.add_paragraph()
    badges = [
        (" 🔄 CPU1 → I/O → CPU2 ", ACCENT_BLUE),
        (" ⏱️ δ Context Switch Cost ", ACCENT_RED),
        (" 🎯 Real-Time Feasibility ", ACCENT_GREEN)
    ]
    for text, col in badges:
        r_pil = p_pills.add_run()
        r_pil.text = text
        r_pil.font.bold = True
        r_pil.font.size = Pt(14)
        r_pil.font.color.rgb = col
        r_sp = p_pills.add_run()
        r_sp.text = "     "

    # Bottom Team Roster
    team_card = create_card(s1, Inches(0.8), Inches(5.35), Inches(11.733), Inches(1.55),
                            bg=CARD_BG, border=RGBColor(38, 50, 72))
    tb_t = s1.shapes.add_textbox(Inches(1.1), Inches(5.45), Inches(11.1), Inches(1.3))
    tf_t = tb_t.text_frame

    pt0 = tf_t.paragraphs[0]
    rt0 = pt0.add_run()
    rt0.text = "👥 Project Team (5 Members):"
    rt0.font.bold = True
    rt0.font.size = Pt(11.5)
    rt0.font.color.rgb = LIGHT_GRAY
    pt0.space_after = Pt(4)

    pt1 = tf_t.add_paragraph()
    rt1 = pt1.add_run()
    rt1.text = "• Sparsh Srivastava (590022251)  |  • Aayushman Kothari (590025499)  |  • Yasharth Jaiswal (590023687)\n• Satvik Arora (590028424)  |  • Yashvardhan Singh (590021777)"
    rt1.font.size = Pt(10.5)
    rt1.font.color.rgb = GRAY_TEXT

    # =========================================================================
    # SLIDE 2: WHY EXISTING SIMULATORS ARE WRONG (Scannable Split: 4 Bullets Each)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Why Traditional Simulators Are Wrong", presenter="Sparsh Srivastava")

    # Left: Traditional
    create_card(s2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.7),
                bg=RGBColor(24, 18, 26), border=ACCENT_RED, lw=2)
    tb_l = s2.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0))
    tf_l = tb_l.text_frame

    pl0 = tf_l.paragraphs[0]
    rl0 = pl0.add_run()
    rl0.text = "❌ Traditional Simulator"
    rl0.font.size = Pt(20)
    rl0.font.bold = True
    rl0.font.color.rgb = ACCENT_RED
    pl0.space_after = Pt(24)

    trad_bullets = [
        "Single CPU Burst Only",
        "Context Switch Overhead (δ = 0)",
        "No Starvation Resolution",
        "No Real-Time Validation"
    ]
    for b in trad_bullets:
        p = tf_l.add_paragraph()
        r = p.add_run()
        r.text = f"• {b}"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = LIGHT_GRAY
        p.space_after = Pt(14)

    # Right: Our Simulator
    create_card(s2, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.7),
                bg=RGBColor(16, 28, 28), border=ACCENT_GREEN, lw=2)
    tb_r = s2.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.0))
    tf_r = tb_r.text_frame

    pr0 = tf_r.paragraphs[0]
    rr0 = pr0.add_run()
    rr0.text = "✅ Our Kernel Simulator"
    rr0.font.size = Pt(20)
    rr0.font.bold = True
    rr0.font.color.rgb = ACCENT_GREEN
    pr0.space_after = Pt(24)

    our_bullets = [
        "CPU1 → I/O → CPU2 Lifecycle",
        "Realistic Context Switch Cost (δ)",
        "Dynamic Priority Aging Engine",
        "RM + EDF Schedulability Analysis"
    ]
    for b in our_bullets:
        p = tf_r.add_paragraph()
        r = p.add_run()
        r.text = f"• {b}"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = WHITE
        p.space_after = Pt(14)

    # Bottom Impact Banner
    banner = create_card(s2, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.8),
                         bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_ban = s2.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.6))
    p_ban = tb_ban.text_frame.paragraphs[0]
    p_ban.alignment = PP_ALIGN.CENTER
    r_ban = p_ban.add_run()
    r_ban.text = "🎯 Core Takeaway: Closer to actual OS kernel behavior than standard classroom tools."
    r_ban.font.size = Pt(13)
    r_ban.font.bold = True
    r_ban.font.color.rgb = ACCENT_BLUE

    # =========================================================================
    # SLIDE 3: PROCESS LIFECYCLE ENGINE (Hero Diagram, 70% Less Text)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Process Lifecycle: 5-State Multi-Burst Engine", presenter="Sparsh Srivastava")

    def draw_state_box(slide, left, top, w, h, name, subtitle, color):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(17, 24, 39)
        card.line.color.rgb = color
        card.line.width = Pt(2.2)

        tb = slide.shapes.add_textbox(left, top + Inches(0.1), w, h - Inches(0.2))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = name
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = color

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(10)
        r2.font.color.rgb = GRAY_TEXT

    # 4 Nodes in Main Row
    draw_state_box(s3, Inches(0.8), Inches(2.3), Inches(2.0), Inches(1.3), "NEW", "Job Queue", ACCENT_BLUE)
    draw_state_box(s3, Inches(3.7), Inches(2.3), Inches(2.3), Inches(1.3), "READY", "RAM Pool", ACCENT_GREEN)
    draw_state_box(s3, Inches(7.0), Inches(2.3), Inches(2.4), Inches(1.3), "RUNNING", "CPU (CPU1/2)", ACCENT_ORANGE)
    draw_state_box(s3, Inches(10.4), Inches(2.3), Inches(2.1), Inches(1.3), "TERMINATED", "Exit Trap", ACCENT_PURPLE)

    # 1 Waiting Node Below
    draw_state_box(s3, Inches(5.3), Inches(4.8), Inches(2.8), Inches(1.3), "WAITING (I/O)", "Device Wait Queue", ACCENT_RED)

    # Forward Arrow Texts
    arrows = [
        (Inches(2.8), Inches(2.55), Inches(0.9), "──▶\nAdmit", ACCENT_BLUE),
        (Inches(6.0), Inches(2.55), Inches(1.0), "──▶\nDispatch", ACCENT_GREEN),
        (Inches(9.4), Inches(2.55), Inches(1.0), "──▶\nComplete", ACCENT_PURPLE)
    ]
    for ax, ay, aw, atxt, acol in arrows:
        tb_a = s3.shapes.add_textbox(ax, ay, aw, Inches(0.8))
        p = tb_a.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = atxt
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = acol

    # Vertical Transitions
    tb_down = s3.shapes.add_textbox(Inches(7.3), Inches(3.75), Inches(2.2), Inches(0.8))
    pd = tb_down.text_frame.paragraphs[0]
    rd = pd.add_run()
    rd.text = "▼ I/O Request (Syscall)"
    rd.font.size = Pt(11)
    rd.font.bold = True
    rd.font.color.rgb = ACCENT_RED

    tb_up = s3.shapes.add_textbox(Inches(3.8), Inches(3.75), Inches(2.2), Inches(0.8))
    pu = tb_up.text_frame.paragraphs[0]
    ru = pu.add_run()
    ru.text = "▲ I/O Done (Interrupt)"
    ru.font.size = Pt(11)
    ru.font.bold = True
    ru.font.color.rgb = ACCENT_GREEN

    # Single Clean Takeaway Pill Below
    create_card(s3, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.8),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_msg = s3.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.6))
    p_m = tb_msg.text_frame.paragraphs[0]
    p_m.alignment = PP_ALIGN.CENTER
    r_m = p_m.add_run()
    r_m.text = "💡 When P1 yields for I/O, CPU is released immediately — allowing Ready Queue tasks to execute."
    r_m.font.size = Pt(12.5)
    r_m.font.bold = True
    r_m.font.color.rgb = ACCENT_BLUE

    # =========================================================================
    # SLIDE 4: ARCHITECTURE PIPELINE (Clean, 2 Bullets Max per Stage)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Inside Our Scheduler: Engine Architecture", presenter="Yashvardhan Singh")

    stages = [
        ("1. Input", ["Multi-burst data", "Priority & periods"], ACCENT_BLUE),
        ("2. Kernel Engine", ["5-State PCB loop", "Device wait queue"], ACCENT_GREEN),
        ("3. Dispatcher", ["FCFS, SJF, SRTF", "Priority & RR"], ACCENT_ORANGE),
        ("4. Context Switch", ["Register save/restore", "Injects δ latency"], ACCENT_RED),
        ("5. Dashboard", ["Interactive Gantt", "Real-time RM/EDF"], ACCENT_PURPLE)
    ]

    cw = Inches(2.15)
    ch = Inches(3.8)
    sp = Inches(0.24)
    sx = Inches(0.8)
    sy = Inches(1.8)

    for idx, (title, b_list, col) in enumerate(stages):
        cur_x = sx + idx * (cw + sp)
        create_card(s4, cur_x, sy, cw, ch, bg=CARD_BG, border=col, lw=2)

        tb = s4.shapes.add_textbox(cur_x + Inches(0.12), sy + Inches(0.2), cw - Inches(0.24), ch - Inches(0.4))
        tf = tb.text_frame
        p_n = tf.paragraphs[0]
        r_n = p_n.add_run()
        r_n.text = f"STAGE {idx+1}"
        r_n.font.size = Pt(10)
        r_n.font.bold = True
        r_n.font.color.rgb = col
        p_n.space_after = Pt(4)

        p_h = tf.add_paragraph()
        r_h = p_h.add_run()
        r_h.text = title
        r_h.font.size = Pt(15)
        r_h.font.bold = True
        r_h.font.color.rgb = WHITE
        p_h.space_after = Pt(16)

        for item in b_list:
            p_i = tf.add_paragraph()
            r_i = p_i.add_run()
            r_i.text = f"• {item}"
            r_i.font.size = Pt(11.5)
            r_i.font.color.rgb = LIGHT_GRAY
            p_i.space_after = Pt(8)

        if idx < len(stages) - 1:
            arr_x = cur_x + cw
            arr_tb = s4.shapes.add_textbox(arr_x, sy + Inches(1.5), sp, Inches(0.6))
            p_a = arr_tb.text_frame.paragraphs[0]
            p_a.alignment = PP_ALIGN.CENTER
            r_a = p_a.add_run()
            r_a.text = "▶"
            r_a.font.size = Pt(14)
            r_a.font.bold = True
            r_a.font.color.rgb = ACCENT_BLUE

    # Bottom Tagline
    create_card(s4, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.9),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_arch = s4.shapes.add_textbox(Inches(1.0), Inches(6.12), Inches(11.3), Inches(0.7))
    p_arch = tb_arch.text_frame.paragraphs[0]
    p_arch.alignment = PP_ALIGN.CENTER
    r_arch = p_arch.add_run()
    r_arch.text = "⚙️ Clean Engineering Separation: Core Simulation (`engine.py`) ➔ Web Dashboard (`app.py`)"
    r_arch.font.size = Pt(13)
    r_arch.font.bold = True
    r_arch.font.color.rgb = ACCENT_BLUE

    # =========================================================================
    # SLIDE 5: CONTEXT SWITCHING COSTS (Huge Visual Number, Memorable)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Context Switching Costs Matter", presenter="Yasharth Jaiswal")

    # Left: RR(q=2)
    create_card(s5, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.7),
                bg=RGBColor(24, 18, 26), border=ACCENT_RED, lw=2.2)
    tb_l = s5.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(5.0), Inches(4.2))
    tf_l = tb_l.text_frame

    pl0 = tf_l.paragraphs[0]
    rl0 = pl0.add_run()
    rl0.text = "Round Robin (q = 2, δ = 1)"
    rl0.font.size = Pt(16)
    rl0.font.bold = True
    rl0.font.color.rgb = ACCENT_RED
    pl0.space_after = Pt(14)

    # Timeline diagram string
    p_pat = tf_l.add_paragraph()
    r_pat = p_pat.add_run()
    r_pat.text = "| P1 | CS | P2 | CS | P3 | CS | P1 | CS |"
    r_pat.font.size = Pt(12)
    r_pat.font.bold = True
    r_pat.font.color.rgb = WHITE
    p_pat.space_after = Pt(20)

    # HUGE RED NUMBER
    p_big = tf_l.add_paragraph()
    r_big = p_big.add_run()
    r_big.text = "27% CPU Lost"
    r_big.font.size = Pt(40)
    r_big.font.bold = True
    r_big.font.color.rgb = ACCENT_RED
    p_big.space_after = Pt(6)

    p_eff = tf_l.add_paragraph()
    r_eff = p_eff.add_run()
    r_eff.text = "Efficiency: η = 73.08%"
    r_eff.font.size = Pt(16)
    r_eff.font.bold = True
    r_eff.font.color.rgb = LIGHT_GRAY

    # Right: RR(q=4)
    create_card(s5, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.7),
                bg=RGBColor(16, 28, 28), border=ACCENT_GREEN, lw=2.2)
    tb_r = s5.shapes.add_textbox(Inches(7.1), Inches(1.75), Inches(5.1), Inches(4.2))
    tf_r = tb_r.text_frame

    pr0 = tf_r.paragraphs[0]
    rr0 = pr0.add_run()
    rr0.text = "Round Robin (q = 4, δ = 1)"
    rr0.font.size = Pt(16)
    rr0.font.bold = True
    rr0.font.color.rgb = ACCENT_GREEN
    pr0.space_after = Pt(14)

    p_pat2 = tf_r.add_paragraph()
    r_pat2 = p_pat2.add_run()
    r_pat2.text = "| P1 ───────── | CS | P2 ───────── | CS |"
    r_pat2.font.size = Pt(12)
    r_pat2.font.bold = True
    r_pat2.font.color.rgb = WHITE
    p_pat2.space_after = Pt(20)

    # HUGE GREEN NUMBER
    p_big2 = tf_r.add_paragraph()
    r_big2 = p_big2.add_run()
    r_big2.text = "+8.7% Restored"
    r_big2.font.size = Pt(40)
    r_big2.font.bold = True
    r_big2.font.color.rgb = ACCENT_GREEN
    p_big2.space_after = Pt(6)

    p_eff2 = tf_r.add_paragraph()
    r_eff2 = p_eff2.add_run()
    r_eff2.text = "Efficiency: η = 81.82%"
    r_eff2.font.size = Pt(16)
    r_eff2.font.bold = True
    r_eff2.font.color.rgb = LIGHT_GRAY

    # Bottom Lesson
    create_card(s5, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.8),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_less = s5.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.6))
    p_l = tb_less.text_frame.paragraphs[0]
    p_l.alignment = PP_ALIGN.CENTER
    r_l = p_l.add_run()
    r_l.text = "⚡ Takeaway: Small quantum improves responsiveness but creates context switch thrashing."
    r_l.font.size = Pt(13)
    r_l.font.bold = True
    r_l.font.color.rgb = ACCENT_BLUE

    # =========================================================================
    # SLIDE 6: STARVATION PROBLEM (Clean Queue Story: P1-P4 stream vs P5)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Starvation Problem & The Aging Solution", presenter="Aayushman Kothari")

    # Left: Starvation Story
    create_card(s6, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.7),
                bg=RGBColor(24, 18, 26), border=ACCENT_RED, lw=2)
    tb_l = s6.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0))
    tf_l = tb_l.text_frame

    pl = tf_l.paragraphs[0]
    rl = pl.add_run()
    rl.text = "⚠️ Without Aging: Indefinite Starvation"
    rl.font.size = Pt(16)
    rl.font.bold = True
    rl.font.color.rgb = ACCENT_RED
    pl.space_after = Pt(20)

    queue_items = [
        ("P1 (Priority = 1) ──▶ Gets CPU", WHITE),
        ("P2 (Priority = 1) ──▶ Gets CPU", WHITE),
        ("P3 (Priority = 1) ──▶ Gets CPU", WHITE),
        ("P4 (Priority = 1) ──▶ Gets CPU", WHITE),
        ("P5 (Priority = 5) ──▶ 😭 Waits Forever!", ACCENT_RED)
    ]
    for q_txt, q_col in queue_items:
        p = tf_l.add_paragraph()
        r = p.add_run()
        r.text = q_txt
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = q_col
        p.space_after = Pt(12)

    # Right: Aging Action
    create_card(s6, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.7),
                bg=RGBColor(16, 28, 28), border=ACCENT_GREEN, lw=2)
    tb_r = s6.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.0))
    tf_r = tb_r.text_frame

    pr = tf_r.paragraphs[0]
    rr = pr.add_run()
    rr.text = "✅ With Dynamic Aging: Guaranteed Execution"
    rr.font.size = Pt(16)
    rr.font.bold = True
    rr.font.color.rgb = ACCENT_GREEN
    pr.space_after = Pt(20)

    aging_story = [
        ("Epoch t=4  ➔  P5 Priority ↑ (5 ➔ 4)", LIGHT_GRAY),
        ("Epoch t=8  ➔  P5 Priority ↑ (4 ➔ 3)", LIGHT_GRAY),
        ("Epoch t=12 ➔  P5 Priority ↑ (3 ➔ 2)", LIGHT_GRAY),
        ("Epoch t=16 ➔  P5 Priority ↑ (2 ➔ 1)", ACCENT_BLUE),
        ("🚀 CPU Assigned ➔ P5 Executes!", ACCENT_GREEN)
    ]
    for a_txt, a_col in aging_story:
        p = tf_r.add_paragraph()
        r = p.add_run()
        r.text = a_txt
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = a_col
        p.space_after = Pt(12)

    # Bottom Tag
    create_card(s6, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.8),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_age = s6.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.6))
    p_a = tb_age.text_frame.paragraphs[0]
    p_a.alignment = PP_ALIGN.CENTER
    r_a = p_a.add_run()
    r_a.text = "🎯 Dynamic aging mathematically guarantees bounded waiting for every process."
    r_a.font.size = Pt(13)
    r_a.font.bold = True
    r_a.font.color.rgb = ACCENT_BLUE

    # =========================================================================
    # SLIDE 7: REAL-TIME SCHEDULING (69.3% and 100% Huge Numbers)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Real-Time Schedulability Bounds", presenter="Satvik Arora")

    # Left Card: RM
    create_card(s7, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.7),
                bg=CARD_BG, border=ACCENT_BLUE, lw=2.2)
    tb_rm = s7.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0))
    tf_rm = tb_rm.text_frame

    p0 = tf_rm.paragraphs[0]
    r0 = p0.add_run()
    r0.text = "Rate Monotonic (RM)"
    r0.font.size = Pt(18)
    r0.font.bold = True
    r0.font.color.rgb = ACCENT_BLUE
    p0.space_after = Pt(4)

    p_sub = tf_rm.add_paragraph()
    r_s = p_sub.add_run()
    r_s.text = "Static Priority Bound"
    r_s.font.size = Pt(13)
    r_s.font.color.rgb = GRAY_TEXT
    p_sub.space_after = Pt(18)

    p_num = tf_rm.add_paragraph()
    r_num = p_num.add_run()
    r_num.text = "69.3%"
    r_num.font.size = Pt(56)
    r_num.font.bold = True
    r_num.font.color.rgb = ACCENT_BLUE
    p_num.space_after = Pt(14)

    p_rm_d = tf_rm.add_paragraph()
    r_rm_d = p_rm_d.add_run()
    r_rm_d.text = "• Static priorities (shorter period = higher priority)\n• U ≤ N(2^(1/N) - 1) ➔ Safe limit is 69.3%"
    r_rm_d.font.size = Pt(12)
    r_rm_d.font.color.rgb = LIGHT_GRAY

    # Right Card: EDF
    create_card(s7, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.7),
                bg=CARD_BG, border=ACCENT_GREEN, lw=2.2)
    tb_edf = s7.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.0))
    tf_edf = tb_edf.text_frame

    p0 = tf_edf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = "Earliest Deadline First (EDF)"
    r0.font.size = Pt(18)
    r0.font.bold = True
    r0.font.color.rgb = ACCENT_GREEN
    p0.space_after = Pt(4)

    p_sub = tf_edf.add_paragraph()
    r_s = p_sub.add_run()
    r_s.text = "Dynamic Priority Bound"
    r_s.font.size = Pt(13)
    r_s.font.color.rgb = GRAY_TEXT
    p_sub.space_after = Pt(18)

    p_num = tf_edf.add_paragraph()
    r_num = p_num.add_run()
    r_num.text = "100.0%"
    r_num.font.size = Pt(56)
    r_num.font.bold = True
    r_num.font.color.rgb = ACCENT_GREEN
    p_num.space_after = Pt(14)

    p_edf_d = tf_edf.add_paragraph()
    r_edf_d = p_edf_d.add_run()
    r_edf_d.text = "• Dynamic priorities based on closest deadline\n• Schedulable if and only if total utilization U ≤ 100%"
    r_edf_d.font.size = Pt(12)
    r_edf_d.font.color.rgb = LIGHT_GRAY

    # Bottom Callout
    create_card(s7, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.8),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_rt_v = s7.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.3), Inches(0.6))
    p_v = tb_rt_v.text_frame.paragraphs[0]
    p_v.alignment = PP_ALIGN.CENTER
    r_v = p_v.add_run()
    r_v.text = "🔬 Automated Check in Our Engine: Test set utilization U = 162% ➔ Flags Infeasibility Alert instantly."
    r_v.font.size = Pt(12.5)
    r_v.font.bold = True
    r_v.font.color.rgb = ACCENT_ORANGE

    # =========================================================================
    # SLIDE 8: FULL SCREEN PROOF (Screenshot Dominates, Zero Paragraphs)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Live Simulator Built By Our Team", presenter="Yashvardhan Singh")

    sim_img_path = os.path.join(assets_dir, "simulator_screenshot.png")
    if os.path.exists(sim_img_path):
        s8.shapes.add_picture(sim_img_path, Inches(0.8), Inches(1.35), width=Inches(11.733), height=Inches(5.75))

    # =========================================================================
    # SLIDE 9: BENCHMARK CHARTS (Charts + Exactly 3 Bullet Takeaways)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Benchmark Results & Empirical Findings", presenter="Aayushman & Yasharth")

    bench_img_path = os.path.join(assets_dir, "benchmark_charts.png")
    if os.path.exists(bench_img_path):
        s9.shapes.add_picture(bench_img_path, Inches(0.8), Inches(1.35), width=Inches(11.733), height=Inches(4.5))

    # Exactly 3 clean bullets below
    create_card(s9, Inches(0.8), Inches(6.0), Inches(11.733), Inches(1.15),
                bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_findings = s9.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.3), Inches(1.05))
    tf_f = tb_findings.text_frame

    pf = tf_f.paragraphs[0]
    rf = pf.add_run()
    rf.text = "Key Findings:"
    rf.font.bold = True
    rf.font.size = Pt(12)
    rf.font.color.rgb = ACCENT_BLUE
    pf.space_after = Pt(2)

    pf1 = tf_f.add_paragraph()
    rf1 = pf1.add_run()
    rf1.text = "• Lowest Initial Response Time ➔ Preemptive Priority with Aging (1.80 units)\n• Highest CPU Utilization Efficiency ➔ FCFS Multi-Burst (85.71%)\n• Worst CPU Efficiency ➔ Round Robin with small quantum q=2 (73.08%)"
    rf1.font.size = Pt(10.5)
    rf1.font.color.rgb = WHITE

    # =========================================================================
    # SLIDE 10: WHAT MAKES THIS DIFFERENT? (4 Cards + Built From Scratch)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "What Makes This Different?", presenter="All 5 Team Members")

    cards = [
        ("🟦 Multi-Burst Model", "Real CPU1 ➔ I/O ➔ CPU2 transitions with device queues.", ACCENT_BLUE),
        ("🟥 Context Switch Cost", "Quantifies dispatcher register save/restore overhead (δ).", ACCENT_RED),
        ("🟩 Dynamic Aging", "Mathematically eliminates low-priority process starvation.", ACCENT_GREEN),
        ("🟪 RM + EDF Analysis", "Validates hard real-time task sets against Liu-Layland bounds.", ACCENT_PURPLE)
    ]

    p_w = Inches(2.72)
    p_h = Inches(3.4)
    p_gap = Inches(0.28)
    p_x = Inches(0.8)
    p_y = Inches(1.6)

    for i, (title, desc, col) in enumerate(cards):
        c_x = p_x + i * (p_w + p_gap)
        create_card(s10, c_x, p_y, p_w, p_h, bg=CARD_BG, border=col, lw=2.2)

        tb = s10.shapes.add_textbox(c_x + Inches(0.12), p_y + Inches(0.2), p_w - Inches(0.24), p_h - Inches(0.4))
        tf = tb.text_frame
        p_t = tf.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.size = Pt(14)
        r_t.font.bold = True
        r_t.font.color.rgb = col
        p_t.space_after = Pt(16)

        p_d = tf.add_paragraph()
        r_d = p_d.add_run()
        r_d.text = desc
        r_d.font.size = Pt(11.5)
        r_d.font.color.rgb = LIGHT_GRAY

    # Bottom Mic-Drop Card
    bot = create_card(s10, Inches(0.8), Inches(5.25), Inches(11.733), Inches(1.8),
                      bg=RGBColor(15, 23, 42), border=ACCENT_ORANGE, lw=2)
    tb_b = s10.shapes.add_textbox(Inches(1.0), Inches(5.35), Inches(11.3), Inches(1.6))
    tf_b = tb_b.text_frame

    pb0 = tf_b.paragraphs[0]
    pb0.alignment = PP_ALIGN.CENTER
    rb0 = pb0.add_run()
    rb0.text = "Built from Scratch using Python, Streamlit & Plotly"
    rb0.font.size = Pt(21)
    rb0.font.bold = True
    rb0.font.color.rgb = ACCENT_ORANGE
    pb0.space_after = Pt(6)

    pb1 = tf_b.add_paragraph()
    pb1.alignment = PP_ALIGN.CENTER
    rb1 = pb1.add_run()
    rb1.text = "Team Members: Sparsh Srivastava | Aayushman Kothari | Yasharth Jaiswal | Satvik Arora | Yashvardhan Singh"
    rb1.font.size = Pt(11)
    rb1.font.color.rgb = WHITE
    pb1.space_after = Pt(4)

    pb2 = tf_b.add_paragraph()
    pb2.alignment = PP_ALIGN.CENTER
    rb2 = pb2.add_run()
    rb2.text = "✨ Thank you! We are open for questions. ✨"
    rb2.font.size = Pt(12)
    rb2.font.bold = True
    rb2.font.color.rgb = ACCENT_GREEN

    # Save
    out_pptx = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\OS_PBL_Presentation_Deck.pptx"
    prs.save(out_pptx)
    print("Human-grade high-impact deck generated successfully:", out_pptx)

if __name__ == "__main__":
    build_human_grade_deck()
