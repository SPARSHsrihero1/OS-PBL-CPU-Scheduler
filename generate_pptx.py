import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_apple_keynote_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 Widescreen
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Modern Engineering Dark Palette
    BG_DARK = RGBColor(15, 23, 42)        # #0F172A Deep Navy
    PANEL_BG = RGBColor(30, 41, 59)       # #1E293B Slate Panel
    PANEL_BORDER = RGBColor(51, 65, 85)   # #334155 Slate Border
    ACCENT_BLUE = RGBColor(59, 130, 246)  # #3B82F6 Blue
    ACCENT_CYAN = RGBColor(6, 182, 212)   # #06B6D4 Cyan
    ACCENT_GREEN = RGBColor(16, 185, 129) # #10B981 Emerald Green
    ACCENT_ORANGE = RGBColor(245, 158, 11)# #F59E0B Warning Orange
    ACCENT_RED = RGBColor(239, 68, 68)    # #EF4444 Red
    ACCENT_PURPLE = RGBColor(168, 85, 247)# #A855F7 Purple
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(148, 163, 184)   # Slate 400
    LIGHT_GRAY = RGBColor(226, 232, 240)  # Slate 200

    assets_dir = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\assets"

    def set_slide_background(slide, color=BG_DARK):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_keynote_header(slide, title, category="OPERATING SYSTEMS PBL", presenter=""):
        # Category Eyebrow
        tb_c = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.0), Inches(0.35))
        pc = tb_c.text_frame.paragraphs[0]
        rc = pc.add_run()
        rc.text = category.upper()
        rc.font.name = "Poppins"
        rc.font.size = Pt(11)
        rc.font.bold = True
        rc.font.color.rgb = ACCENT_CYAN

        # Title
        tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(9.2), Inches(0.75))
        pt = tb_t.text_frame.paragraphs[0]
        rt = pt.add_run()
        rt.text = title
        rt.font.name = "Montserrat"
        rt.font.size = Pt(28)
        rt.font.bold = True
        rt.font.color.rgb = WHITE

        # Presenter Tag
        if presenter:
            tb_p = slide.shapes.add_textbox(Inches(9.2), Inches(0.55), Inches(3.3), Inches(0.55))
            pp = tb_p.text_frame.paragraphs[0]
            pp.alignment = PP_ALIGN.RIGHT
            rp = pp.add_run()
            rp.text = f"🎙️ {presenter}"
            rp.font.name = "Inter"
            rp.font.size = Pt(12)
            rp.font.bold = True
            rp.font.color.rgb = ACCENT_BLUE

    def create_card(slide, left, top, width, height, bg=PANEL_BG, border=PANEL_BORDER, lw=1.5):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = border
        card.line.width = Pt(lw)
        return card

    # =========================================================================
    # SLIDE 1: BIG IMPACT HERO (Keynote Style, Aggressive Whitespace)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, BG_DARK)

    # Hero Center Box
    tb_hero = s1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(3.6))
    tf1 = tb_hero.text_frame

    p_badge = tf1.paragraphs[0]
    rb = p_badge.add_run()
    rb.text = "OPERATING SYSTEMS PBL SHOWCASE"
    rb.font.name = "Poppins"
    rb.font.size = Pt(13)
    rb.font.bold = True
    rb.font.color.rgb = ACCENT_ORANGE
    p_badge.space_after = Pt(14)

    p_title = tf1.add_paragraph()
    rt = p_title.add_run()
    rt.text = "CPU Scheduling Beyond Textbooks"
    rt.font.name = "Montserrat"
    rt.font.size = Pt(44)
    rt.font.bold = True
    rt.font.color.rgb = WHITE
    p_title.space_after = Pt(8)

    p_sub = tf1.add_paragraph()
    rs = p_sub.add_run()
    rs.text = "Modeling Real Operating System Kernel Mechanics"
    rs.font.name = "Inter"
    rs.font.size = Pt(22)
    rs.font.bold = True
    rs.font.color.rgb = ACCENT_CYAN
    p_sub.space_after = Pt(28)

    # 3 High-Impact Visual Pills
    p_pills = tf1.add_paragraph()
    pills = [
        (" 🔄 Multi-Burst Lifecycle ", ACCENT_BLUE),
        (" ⏱️ Context Switch Overhead (δ) ", ACCENT_ORANGE),
        (" 🎯 Real-Time Feasibility (RM/EDF) ", ACCENT_GREEN)
    ]
    for txt, col in pills:
        rp = p_pills.add_run()
        rp.text = txt
        rp.font.name = "Inter"
        rp.font.bold = True
        rp.font.size = Pt(14)
        rp.font.color.rgb = col
        rsp = p_pills.add_run()
        rsp.text = "     "

    # Team Footer (Clean, Subtle)
    team_card = create_card(s1, Inches(1.0), Inches(5.4), Inches(11.333), Inches(1.4),
                            bg=PANEL_BG, border=PANEL_BORDER, lw=1.2)
    tb_team = s1.shapes.add_textbox(Inches(1.2), Inches(5.5), Inches(10.9), Inches(1.2))
    tft = tb_team.text_frame

    pt0 = tft.paragraphs[0]
    rt0 = pt0.add_run()
    rt0.text = "TEAM 7 | B.TECH COMPUTER SCIENCE & ENGINEERING"
    rt0.font.name = "Poppins"
    rt0.font.size = Pt(11)
    rt0.font.bold = True
    rt0.font.color.rgb = ACCENT_CYAN
    pt0.space_after = Pt(4)

    pt1 = tft.add_paragraph()
    rt1 = pt1.add_run()
    rt1.text = "Sparsh Srivastava (590022251)  •  Aayushman Kothari (590025499)  •  Yasharth Jaiswal (590023687)\nSatvik Arora (590028424)  •  Yashvardhan Singh (590021777)"
    rt1.font.name = "Inter"
    rt1.font.size = Pt(12)
    rt1.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 2: PROBLEM COMPARISON (Two Clean Columns, Zero Paragraphs)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_keynote_header(s2, "Traditional Simulator vs Our Simulator", presenter="Sparsh Srivastava")

    # Left Card: Traditional
    create_card(s2, Inches(0.8), Inches(1.6), Inches(5.6), Inches(4.5),
                bg=RGBColor(24, 20, 28), border=ACCENT_RED, lw=2)
    tb_l = s2.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(3.8))
    tfl = tb_l.text_frame

    pl0 = tfl.paragraphs[0]
    rl0 = pl0.add_run()
    rl0.text = "❌ Traditional Simulator"
    rl0.font.name = "Montserrat"
    rl0.font.size = Pt(22)
    rl0.font.bold = True
    rl0.font.color.rgb = ACCENT_RED
    pl0.space_after = Pt(24)

    trad_points = [
        "Single CPU burst only",
        "Context switch cost (δ = 0)",
        "No dynamic aging (starvation)",
        "No real-time deadline validation"
    ]
    for pt_text in trad_points:
        p = tfl.add_paragraph()
        r = p.add_run()
        r.text = f"•  {pt_text}"
        r.font.name = "Inter"
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = LIGHT_GRAY
        p.space_after = Pt(14)

    # Right Card: Our Simulator
    create_card(s2, Inches(6.8), Inches(1.6), Inches(5.7), Inches(4.5),
                bg=RGBColor(16, 32, 30), border=ACCENT_GREEN, lw=2)
    tb_r = s2.shapes.add_textbox(Inches(7.1), Inches(1.9), Inches(5.1), Inches(3.8))
    tfr = tb_r.text_frame

    pr0 = tfr.paragraphs[0]
    rr0 = pr0.add_run()
    rr0.text = "✅ Our Kernel Simulator"
    rr0.font.name = "Montserrat"
    rr0.font.size = Pt(22)
    rr0.font.bold = True
    rr0.font.color.rgb = ACCENT_GREEN
    pr0.space_after = Pt(24)

    our_points = [
        "CPU1 → I/O → CPU2 multi-burst",
        "Realistic context switch latency (δ)",
        "Dynamic priority aging engine",
        "RM + EDF schedulability checks"
    ]
    for pt_text in our_points:
        p = tfr.add_paragraph()
        r = p.add_run()
        r.text = f"•  {pt_text}"
        r.font.name = "Inter"
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = WHITE
        p.space_after = Pt(14)

    # Bottom Single-Idea Banner
    create_card(s2, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_bot2 = s2.shapes.add_textbox(Inches(1.0), Inches(6.43), Inches(11.3), Inches(0.6))
    p_b2 = tb_bot2.text_frame.paragraphs[0]
    p_b2.alignment = PP_ALIGN.CENTER
    r_b2 = p_b2.add_run()
    r_b2.text = "🎯 Core Difference: Built to reflect real kernel behavior, not textbook simplifications."
    r_b2.font.name = "Inter"
    r_b2.font.size = Pt(14)
    r_b2.font.bold = True
    r_b2.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 3: PROCESS LIFECYCLE (State-Transition Diagram Dominates 70%)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_keynote_header(s3, "Process Lifecycle: Multi-Burst State Machine", presenter="Sparsh Srivastava")

    def draw_state_box(slide, left, top, w, h, name, subtitle, color):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = PANEL_BG
        card.line.color.rgb = color
        card.line.width = Pt(2.5)

        tb = slide.shapes.add_textbox(left, top + Inches(0.12), w, h - Inches(0.24))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = name
        r.font.name = "Montserrat"
        r.font.size = Pt(17)
        r.font.bold = True
        r.font.color.rgb = color

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.name = "Inter"
        r2.font.size = Pt(11)
        r2.font.color.rgb = GRAY_TEXT

    # Diagram Nodes (Main Row)
    draw_state_box(s3, Inches(0.8), Inches(2.2), Inches(2.0), Inches(1.4), "NEW", "Job Queue", ACCENT_BLUE)
    draw_state_box(s3, Inches(3.7), Inches(2.2), Inches(2.3), Inches(1.4), "READY", "RAM Pool", ACCENT_GREEN)
    draw_state_box(s3, Inches(7.0), Inches(2.2), Inches(2.4), Inches(1.4), "RUNNING", "CPU Allocation", ACCENT_ORANGE)
    draw_state_box(s3, Inches(10.4), Inches(2.2), Inches(2.1), Inches(1.4), "TERMINATED", "PCB Reclaimed", ACCENT_PURPLE)

    # Waiting Node (Below)
    draw_state_box(s3, Inches(5.3), Inches(4.7), Inches(2.8), Inches(1.4), "WAITING (I/O)", "Device Wait Queue", ACCENT_RED)

    # Transition Arrows
    transitions = [
        (Inches(2.8), Inches(2.5), Inches(0.9), "──▶\nAdmit", ACCENT_BLUE),
        (Inches(6.0), Inches(2.5), Inches(1.0), "──▶\nDispatch", ACCENT_GREEN),
        (Inches(9.4), Inches(2.5), Inches(1.0), "──▶\nExit", ACCENT_PURPLE)
    ]
    for ax, ay, aw, atxt, acol in transitions:
        tb_a = s3.shapes.add_textbox(ax, ay, aw, Inches(0.8))
        p = tb_a.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = atxt
        r.font.name = "Inter"
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = acol

    # Vertical Transitions
    tb_down = s3.shapes.add_textbox(Inches(7.3), Inches(3.7), Inches(2.4), Inches(0.7))
    pd = tb_down.text_frame.paragraphs[0]
    rd = pd.add_run()
    rd.text = "▼ Syscall Trap (I/O Req)"
    rd.font.name = "Inter"
    rd.font.size = Pt(12)
    rd.font.bold = True
    rd.font.color.rgb = ACCENT_RED

    tb_up = s3.shapes.add_textbox(Inches(3.7), Inches(3.7), Inches(2.4), Inches(0.7))
    pu = tb_up.text_frame.paragraphs[0]
    ru = pu.add_run()
    ru.text = "▲ I/O Done (Interrupt)"
    ru.font.name = "Inter"
    ru.font.size = Pt(12)
    ru.font.bold = True
    ru.font.color.rgb = ACCENT_GREEN

    # Only 3 Clean Takeaways at Bottom
    create_card(s3, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_take3 = s3.shapes.add_textbox(Inches(1.0), Inches(6.43), Inches(11.3), Inches(0.6))
    p_t3 = tb_take3.text_frame.paragraphs[0]
    p_t3.alignment = PP_ALIGN.CENTER
    r_t3 = p_t3.add_run()
    r_t3.text = "💡 Key Takeaway: CPU is unblocked during I/O — allowing Ready Queue tasks to interleave execution."
    r_t3.font.name = "Inter"
    r_t3.font.size = Pt(13)
    r_t3.font.bold = True
    r_t3.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 4: SYSTEM ARCHITECTURE (Modern Pipeline Diagram Dominates)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_keynote_header(s4, "System Architecture Pipeline", presenter="Yashvardhan Singh")

    pipeline = [
        ("User Input", "Process Config\nBurst Vectors\nPeriods", ACCENT_BLUE),
        ("Streamlit UI", "Dynamic Editor\nParameter Sliders\nInstant Feedback", ACCENT_CYAN),
        ("Scheduler Engine", "Discrete Tick Loop\n5-State PCB Manager\nDevice Queue", ACCENT_GREEN),
        ("Dispatcher", "FCFS / SJF / SRTF\nPriority & Aging\nRound Robin (q)", ACCENT_ORANGE),
        ("Context Switch", "Register Save/Restore\nInjects δ Delay\nTracks Efficiency", ACCENT_RED),
        ("Visualization", "Plotly Gantt Chart\nPCB State Table\nRM/EDF Analytics", ACCENT_PURPLE)
    ]

    card_w = Inches(1.8)
    card_h = Inches(4.2)
    spacing = Inches(0.18)
    start_x = Inches(0.8)
    start_y = Inches(1.7)

    for idx, (title, content, col) in enumerate(pipeline):
        cur_x = start_x + idx * (card_w + spacing)
        create_card(s4, cur_x, start_y, card_w, card_h, bg=PANEL_BG, border=col, lw=2.2)

        tb = s4.shapes.add_textbox(cur_x + Inches(0.08), start_y + Inches(0.2), card_w - Inches(0.16), card_h - Inches(0.4))
        tf = tb.text_frame

        p_num = tf.paragraphs[0]
        rn = p_num.add_run()
        rn.text = f"0{idx+1}"
        rn.font.name = "Montserrat"
        rn.font.size = Pt(11)
        rn.font.bold = True
        rn.font.color.rgb = col
        p_num.space_after = Pt(4)

        p_h = tf.add_paragraph()
        rh = p_h.add_run()
        rh.text = title
        rh.font.name = "Montserrat"
        rh.font.size = Pt(15)
        rh.font.bold = True
        rh.font.color.rgb = WHITE
        p_h.space_after = Pt(18)

        for line in content.split('\n'):
            pl = tf.add_paragraph()
            rl = pl.add_run()
            rl.text = f"• {line}"
            rl.font.name = "Inter"
            rl.font.size = Pt(11)
            rl.font.color.rgb = LIGHT_GRAY
            pl.space_after = Pt(6)

        if idx < len(pipeline) - 1:
            arr_x = cur_x + card_w
            arr_tb = s4.shapes.add_textbox(arr_x, start_y + Inches(1.8), spacing, Inches(0.6))
            pa = arr_tb.text_frame.paragraphs[0]
            pa.alignment = PP_ALIGN.CENTER
            ra = pa.add_run()
            ra.text = "▶"
            ra.font.name = "Inter"
            ra.font.size = Pt(13)
            ra.font.bold = True
            ra.font.color.rgb = ACCENT_BLUE

    # Single Bottom Banner
    create_card(s4, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_arch_bot = s4.shapes.add_textbox(Inches(1.0), Inches(6.33), Inches(11.3), Inches(0.6))
    pa_bot = tb_arch_bot.text_frame.paragraphs[0]
    pa_bot.alignment = PP_ALIGN.CENTER
    ra_bot = pa_bot.add_run()
    ra_bot.text = "⚙️ Clean Engineering Separation: Core Simulation (`engine.py`) ➔ Interactive GUI (`app.py`)"
    ra_bot.font.name = "Inter"
    ra_bot.font.size = Pt(13)
    ra_bot.font.bold = True
    ra_bot.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 5: CONTEXT SWITCHING (Large 73% and 82% Dominate Slide)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_keynote_header(s5, "Context Switching: Smaller Quantum Causes Efficiency Loss", presenter="Yasharth Jaiswal")

    # Left Card: q = 2
    create_card(s5, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.5),
                bg=RGBColor(24, 20, 28), border=ACCENT_RED, lw=2.2)
    tb_l5 = s5.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0))
    tf_l5 = tb_l5.text_frame

    p_qh1 = tf_l5.paragraphs[0]
    rq1 = p_qh1.add_run()
    rq1.text = "Round Robin (q = 2, δ = 1)"
    rq1.font.name = "Montserrat"
    rq1.font.size = Pt(18)
    rq1.font.bold = True
    rq1.font.color.rgb = ACCENT_RED
    p_qh1.space_after = Pt(10)

    # Timeline string illustration
    p_tim1 = tf_l5.add_paragraph()
    rt1 = p_tim1.add_run()
    rt1.text = "| P1 | CS | P2 | CS | P3 | CS | P1 | CS |"
    rt1.font.name = "Inter"
    rt1.font.size = Pt(13)
    rt1.font.bold = True
    rt1.font.color.rgb = WHITE
    p_tim1.space_after = Pt(16)

    # GIANT NUMBER: 73%
    p_num1 = tf_l5.add_paragraph()
    rn1 = p_num1.add_run()
    rn1.text = "73%"
    rn1.font.name = "Montserrat"
    rn1.font.size = Pt(64)
    rn1.font.bold = True
    rn1.font.color.rgb = ACCENT_RED
    p_num1.space_after = Pt(4)

    p_lbl1 = tf_l5.add_paragraph()
    rl1 = p_lbl1.add_run()
    rl1.text = "CPU Efficiency (η)  •  27% CPU Lost to Preemption!"
    rl1.font.name = "Inter"
    rl1.font.size = Pt(13)
    rl1.font.bold = True
    rl1.font.color.rgb = LIGHT_GRAY

    # Right Card: q = 4
    create_card(s5, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.5),
                bg=RGBColor(16, 32, 30), border=ACCENT_GREEN, lw=2.2)
    tb_r5 = s5.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.0))
    tf_r5 = tb_r5.text_frame

    p_qh2 = tf_r5.paragraphs[0]
    rq2 = p_qh2.add_run()
    rq2.text = "Round Robin (q = 4, δ = 1)"
    rq2.font.name = "Montserrat"
    rq2.font.size = Pt(18)
    rq2.font.bold = True
    rq2.font.color.rgb = ACCENT_GREEN
    p_qh2.space_after = Pt(10)

    p_tim2 = tf_r5.add_paragraph()
    rt2 = p_tim2.add_run()
    rt2.text = "| P1 ───────── | CS | P2 ───────── | CS |"
    rt2.font.name = "Inter"
    rt2.font.size = Pt(13)
    rt2.font.bold = True
    rt2.font.color.rgb = WHITE
    p_tim2.space_after = Pt(16)

    # GIANT NUMBER: 82%
    p_num2 = tf_r5.add_paragraph()
    rn2 = p_num2.add_run()
    rn2.text = "82%"
    rn2.font.name = "Montserrat"
    rn2.font.size = Pt(64)
    rn2.font.bold = True
    rn2.font.color.rgb = ACCENT_GREEN
    p_num2.space_after = Pt(4)

    p_lbl2 = tf_r5.add_paragraph()
    rl2 = p_lbl2.add_run()
    rl2.text = "CPU Efficiency (η)  •  +9% Restored without latency hit!"
    rl2.font.name = "Inter"
    rl2.font.size = Pt(13)
    rl2.font.bold = True
    rl2.font.color.rgb = LIGHT_GRAY

    # Bottom Single Lesson
    create_card(s5, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_bot5 = s5.shapes.add_textbox(Inches(1.0), Inches(6.43), Inches(11.3), Inches(0.6))
    p_b5 = tb_bot5.text_frame.paragraphs[0]
    p_b5.alignment = PP_ALIGN.CENTER
    r_b5 = p_b5.add_run()
    r_b5.text = "⚡ Engineering Principle: Quantum must be large relative to context switch overhead (q >> δ)."
    r_b5.font.name = "Inter"
    r_b5.font.size = Pt(14)
    r_b5.font.bold = True
    r_b5.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 6: DYNAMIC AGING (Visual Story: Waiting Process → Priority ↑ → CPU)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_keynote_header(s6, "Dynamic Aging: Eliminating Process Starvation", presenter="Aayushman Kothari")

    # 3 Story Cards
    card_w = Inches(3.64)
    card_h = Inches(4.5)
    gap = Inches(0.4)
    y_pos = Inches(1.6)

    # Step 1: Waiting Process
    create_card(s6, Inches(0.8), y_pos, card_w, card_h, bg=RGBColor(24, 20, 28), border=ACCENT_RED, lw=2)
    tb_s1 = s6.shapes.add_textbox(Inches(1.0), y_pos + Inches(0.2), card_w - Inches(0.4), card_h - Inches(0.4))
    tfs1 = tb_s1.text_frame

    ps1_h = tfs1.paragraphs[0]
    rs1_h = ps1_h.add_run()
    rs1_h.text = "1. Waiting Process"
    rs1_h.font.name = "Montserrat"
    rs1_h.font.size = Pt(18)
    rs1_h.font.bold = True
    rs1_h.font.color.rgb = ACCENT_RED
    ps1_h.space_after = Pt(20)

    p_queue = tfs1.add_paragraph()
    rq = p_queue.add_run()
    rq.text = "Ready Queue:\n\n• P1 (Priority 1) ➔ Dispatched\n• P2 (Priority 1) ➔ Dispatched\n• P3 (Priority 1) ➔ Dispatched\n\n• P5 (Priority 5) ➔ 😭 Starved!"
    rq.font.name = "Inter"
    rq.font.size = Pt(13)
    rq.font.bold = True
    rq.font.color.rgb = LIGHT_GRAY

    # Step 2: Priority Increases
    create_card(s6, Inches(0.8) + card_w + gap, y_pos, card_w, card_h, bg=PANEL_BG, border=ACCENT_ORANGE, lw=2)
    tb_s2 = s6.shapes.add_textbox(Inches(1.0) + card_w + gap, y_pos + Inches(0.2), card_w - Inches(0.4), card_h - Inches(0.4))
    tfs2 = tb_s2.text_frame

    ps2_h = tfs2.paragraphs[0]
    rs2_h = ps2_h.add_run()
    rs2_h.text = "2. Priority Increases"
    rs2_h.font.name = "Montserrat"
    rs2_h.font.size = Pt(18)
    rs2_h.font.bold = True
    rs2_h.font.color.rgb = ACCENT_ORANGE
    ps2_h.space_after = Pt(20)

    p_aging = tfs2.add_paragraph()
    ra = p_aging.add_run()
    ra.text = "Aging Escalation:\n\n• t = 4  ➔  P5 Priority ↑ (4)\n• t = 8  ➔  P5 Priority ↑ (3)\n• t = 12 ➔  P5 Priority ↑ (2)\n• t = 16 ➔  P5 Priority ↑ (1)"
    ra.font.name = "Inter"
    ra.font.size = Pt(13)
    ra.font.bold = True
    ra.font.color.rgb = WHITE

    # Step 3: Receives CPU
    create_card(s6, Inches(0.8) + 2*(card_w + gap), y_pos, card_w, card_h, bg=RGBColor(16, 32, 30), border=ACCENT_GREEN, lw=2)
    tb_s3 = s6.shapes.add_textbox(Inches(1.0) + 2*(card_w + gap), y_pos + Inches(0.2), card_w - Inches(0.4), card_h - Inches(0.4))
    tfs3 = tb_s3.text_frame

    ps3_h = tfs3.paragraphs[0]
    rs3_h = ps3_h.add_run()
    rs3_h.text = "3. Receives CPU"
    rs3_h.font.name = "Montserrat"
    rs3_h.font.size = Pt(18)
    rs3_h.font.bold = True
    rs3_h.font.color.rgb = ACCENT_GREEN
    ps3_h.space_after = Pt(20)

    p_exec = tfs3.add_paragraph()
    re = p_exec.add_run()
    re.text = "Execution Granted:\n\n• Highest Priority Achieved\n• Dispatched to CPU\n• Finishes execution\n\n✅ Starvation Eliminated!"
    re.font.name = "Inter"
    re.font.size = Pt(13)
    re.font.bold = True
    re.font.color.rgb = WHITE

    # Bottom Tag
    create_card(s6, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_bot6 = s6.shapes.add_textbox(Inches(1.0), Inches(6.43), Inches(11.3), Inches(0.6))
    p_b6 = tb_bot6.text_frame.paragraphs[0]
    p_b6.alignment = PP_ALIGN.CENTER
    r_b6 = p_b6.add_run()
    r_b6.text = "🎯 Dynamic aging mathematically guarantees bounded waiting for every process."
    r_b6.font.name = "Inter"
    r_b6.font.size = Pt(14)
    r_b6.font.bold = True
    r_b6.font.color.rgb = ACCENT_CYAN

    # =========================================================================
    # SLIDE 7: RM VS EDF (Large 69.3% vs 100% Numbers are Visual Focus)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_keynote_header(s7, "Real-Time Schedulability: RM vs EDF", presenter="Satvik Arora")

    # Left: RM
    create_card(s7, Inches(0.8), Inches(1.5), Inches(5.6), Inches(4.5),
                bg=PANEL_BG, border=ACCENT_BLUE, lw=2.5)
    tb_rm = s7.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0))
    tf_rm = tb_rm.text_frame

    p_rm0 = tf_rm.paragraphs[0]
    r_rm0 = p_rm0.add_run()
    r_rm0.text = "Rate Monotonic (RM)"
    r_rm0.font.name = "Montserrat"
    r_rm0.font.size = Pt(20)
    r_rm0.font.bold = True
    r_rm0.font.color.rgb = ACCENT_BLUE
    p_rm0.space_after = Pt(2)

    p_rm_sub = tf_rm.add_paragraph()
    r_rms = p_rm_sub.add_run()
    r_rms.text = "Static Priority Utilization Bound"
    r_rms.font.name = "Inter"
    r_rms.font.size = Pt(13)
    r_rms.font.color.rgb = GRAY_TEXT
    p_rm_sub.space_after = Pt(14)

    # GIANT NUMBER: 69.3%
    p_num_rm = tf_rm.add_paragraph()
    r_num_rm = p_num_rm.add_run()
    r_num_rm.text = "69.3%"
    r_num_rm.font.name = "Montserrat"
    r_num_rm.font.size = Pt(64)
    r_num_rm.font.bold = True
    r_num_rm.font.color.rgb = ACCENT_BLUE
    p_num_rm.space_after = Pt(10)

    p_rm_pt = tf_rm.add_paragraph()
    r_rpt = p_rm_pt.add_run()
    r_rpt.text = "• Priority inversely proportional to period\n• Mathematical guarantee if U ≤ 69.3%"
    r_rpt.font.name = "Inter"
    r_rpt.font.size = Pt(13)
    r_rpt.font.color.rgb = LIGHT_GRAY

    # Right: EDF
    create_card(s7, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.5),
                bg=PANEL_BG, border=ACCENT_GREEN, lw=2.5)
    tb_edf = s7.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.0))
    tf_edf = tb_edf.text_frame

    p_edf0 = tf_edf.paragraphs[0]
    r_edf0 = p_edf0.add_run()
    r_edf0.text = "Earliest Deadline First (EDF)"
    r_edf0.font.name = "Montserrat"
    r_edf0.font.size = Pt(20)
    r_edf0.font.bold = True
    r_edf0.font.color.rgb = ACCENT_GREEN
    p_edf0.space_after = Pt(2)

    p_edf_sub = tf_edf.add_paragraph()
    r_edfs = p_edf_sub.add_run()
    r_edfs.text = "Dynamic Priority Utilization Bound"
    r_edfs.font.name = "Inter"
    r_edfs.font.size = Pt(13)
    r_edfs.font.color.rgb = GRAY_TEXT
    p_edf_sub.space_after = Pt(14)

    # GIANT NUMBER: 100%
    p_num_edf = tf_edf.add_paragraph()
    r_num_edf = p_num_edf.add_run()
    r_num_edf.text = "100%"
    r_num_edf.font.name = "Montserrat"
    r_num_edf.font.size = Pt(64)
    r_num_edf.font.bold = True
    r_num_edf.font.color.rgb = ACCENT_GREEN
    p_num_edf.space_after = Pt(10)

    p_edf_pt = tf_edf.add_paragraph()
    r_ept = p_edf_pt.add_run()
    r_ept.text = "• Priority based on nearest absolute deadline\n• Fully schedulable up to 100% CPU capacity"
    r_ept.font.name = "Inter"
    r_ept.font.size = Pt(13)
    r_ept.font.color.rgb = LIGHT_GRAY

    # Bottom Single Lesson
    create_card(s7, Inches(0.8), Inches(6.35), Inches(11.733), Inches(0.75),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_bot7 = s7.shapes.add_textbox(Inches(1.0), Inches(6.43), Inches(11.3), Inches(0.6))
    p_b7 = tb_bot7.text_frame.paragraphs[0]
    p_b7.alignment = PP_ALIGN.CENTER
    r_b7 = p_b7.add_run()
    r_b7.text = "🔬 Our Engine: Test set U = 162% ➔ Instant automated Infeasibility Alert."
    r_b7.font.name = "Inter"
    r_b7.font.size = Pt(13)
    r_b7.font.bold = True
    r_b7.font.color.rgb = ACCENT_ORANGE

    # =========================================================================
    # SLIDE 8: PRODUCT DEMONSTRATION (Screenshot Occupies 85% of Slide Area)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_keynote_header(s8, "Product Demonstration: Live Kernel Simulator", presenter="Yashvardhan Singh")

    sim_img_path = os.path.join(assets_dir, "simulator_screenshot.png")
    if os.path.exists(sim_img_path):
        # Occupies 85% of slide area with sleek border card
        s8.shapes.add_picture(sim_img_path, Inches(0.8), Inches(1.35), width=Inches(11.733), height=Inches(5.75))

    # =========================================================================
    # SLIDE 9: BENCHMARK RESULTS (Charts + 3 Bullet Highlights)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_keynote_header(s9, "Benchmark Results & Empirical Trade-Offs", presenter="Aayushman & Yasharth")

    bench_img_path = os.path.join(assets_dir, "benchmark_charts.png")
    if os.path.exists(bench_img_path):
        s9.shapes.add_picture(bench_img_path, Inches(0.8), Inches(1.35), width=Inches(11.733), height=Inches(4.5))

    # Exactly 3 clean bullets below
    create_card(s9, Inches(0.8), Inches(6.05), Inches(11.733), Inches(1.1),
                bg=PANEL_BG, border=PANEL_BORDER)
    tb_b9 = s9.shapes.add_textbox(Inches(1.0), Inches(6.12), Inches(11.3), Inches(0.95))
    tf_b9 = tb_b9.text_frame

    p_f0 = tf_b9.paragraphs[0]
    rf0 = p_f0.add_run()
    rf0.text = "Key Highlights:"
    rf0.font.name = "Poppins"
    rf0.font.size = Pt(11)
    rf0.font.bold = True
    rf0.font.color.rgb = ACCENT_CYAN
    p_f0.space_after = Pt(2)

    p_f1 = tf_b9.add_paragraph()
    rf1 = p_f1.add_run()
    rf1.text = "• Best Response Time ➔ Preemptive Priority with Aging (1.80 units)\n• Best CPU Efficiency ➔ FCFS Multi-Burst (85.71%)\n• Trade-off ➔ Small Round Robin quantum (q=2) suffers heavy preemption overhead (73.08%)"
    rf1.font.name = "Inter"
    rf1.font.size = Pt(11)
    rf1.font.bold = True
    rf1.font.color.rgb = WHITE

    # =========================================================================
    # SLIDE 10: PROJECT CONTRIBUTION (4 Visual Cards + Single Closing Statement)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_keynote_header(s10, "Project Contribution", presenter="All 5 Members")

    cards = [
        ("Multi-Burst Lifecycle", "Real CPU1 ➔ I/O ➔ CPU2 transitions with device wait queue management.", ACCENT_BLUE),
        ("Context Switching Analysis", "Explicit quantification of dispatcher register save/restore overhead (δ).", ACCENT_ORANGE),
        ("Dynamic Aging", "Mathematical guarantee against process starvation in priority queues.", ACCENT_GREEN),
        ("Real-Time Scheduling", "Automated Liu-Layland schedulability analysis (RM ≤ 69.3%, EDF ≤ 100%).", ACCENT_PURPLE)
    ]

    p_w = Inches(2.72)
    p_h = Inches(3.4)
    p_gap = Inches(0.28)
    p_x = Inches(0.8)
    p_y = Inches(1.6)

    for i, (title, desc, col) in enumerate(cards):
        c_x = p_x + i * (p_w + p_gap)
        create_card(s10, c_x, p_y, p_w, p_h, bg=PANEL_BG, border=col, lw=2.2)

        tb = s10.shapes.add_textbox(c_x + Inches(0.12), p_y + Inches(0.2), p_w - Inches(0.24), p_h - Inches(0.4))
        tf = tb.text_frame

        p_ct = tf.paragraphs[0]
        r_ct = p_ct.add_run()
        r_ct.text = title
        r_ct.font.name = "Montserrat"
        r_ct.font.size = Pt(15)
        r_ct.font.bold = True
        r_ct.font.color.rgb = col
        p_ct.space_after = Pt(16)

        p_cd = tf.add_paragraph()
        r_cd = p_cd.add_run()
        r_cd.text = desc
        r_cd.font.name = "Inter"
        r_cd.font.size = Pt(12)
        r_cd.font.color.rgb = LIGHT_GRAY

    # One Powerful Closing Statement (No Paragraphs)
    create_card(s10, Inches(0.8), Inches(5.3), Inches(11.733), Inches(1.7),
                bg=PANEL_BG, border=ACCENT_CYAN, lw=2.2)
    tb_close = s10.shapes.add_textbox(Inches(1.0), Inches(5.42), Inches(11.3), Inches(1.4))
    tfc = tb_close.text_frame

    pc0 = tfc.paragraphs[0]
    pc0.alignment = PP_ALIGN.CENTER
    rc0 = pc0.add_run()
    rc0.text = "Closer to Real Operating System Kernels Than Traditional Simulators"
    rc0.font.name = "Montserrat"
    rc0.font.size = Pt(21)
    rc0.font.bold = True
    rc0.font.color.rgb = ACCENT_CYAN
    pc0.space_after = Pt(6)

    pc1 = tfc.add_paragraph()
    pc1.alignment = PP_ALIGN.CENTER
    rc1 = pc1.add_run()
    rc1.text = "Built from Scratch in Python, Streamlit & Plotly  •  Team 7"
    rc1.font.name = "Inter"
    rc1.font.size = Pt(13)
    rc1.font.bold = True
    rc1.font.color.rgb = WHITE
    pc1.space_after = Pt(4)

    pc2 = tfc.add_paragraph()
    pc2.alignment = PP_ALIGN.CENTER
    rc2 = pc2.add_run()
    rc2.text = "✨ Thank you. We are ready for Viva Voce examination! ✨"
    rc2.font.name = "Inter"
    rc2.font.size = Pt(12)
    rc2.font.bold = True
    rc2.font.color.rgb = ACCENT_GREEN

    # Save
    out_pptx = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\OS_PBL_Presentation_Deck.pptx"
    prs.save(out_pptx)
    print("Apple Keynote / MIT Showcase style deck built successfully:", out_pptx)

if __name__ == "__main__":
    build_apple_keynote_deck()
