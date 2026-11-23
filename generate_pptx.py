import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_top1_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 Widescreen
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Executive Modern Dark Theme Palette
    BG_DARK = RGBColor(11, 15, 25)        # #0B0F19 Dark Navy Slate
    PANEL_DARK = RGBColor(17, 24, 39)     # #111827 Dark Card Slate
    PANEL_BORDER = RGBColor(31, 41, 55)   # #1F2937 Border
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

    def add_slide_header(slide, title, category="OPERATING SYSTEMS PBL ACTIVITY | 20% IA WEIGHTAGE", presenter=""):
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.5), Inches(0.35))
        p_c = tb_cat.text_frame.paragraphs[0]
        r_c = p_c.add_run()
        r_c.text = category.upper()
        r_c.font.size = Pt(10.5)
        r_c.font.bold = True
        r_c.font.color.rgb = ACCENT_ORANGE

        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(9.2), Inches(0.75))
        p_t = tb_title.text_frame.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.size = Pt(21)
        r_t.font.bold = True
        r_t.font.color.rgb = WHITE

        if presenter:
            tb_p = slide.shapes.add_textbox(Inches(9.2), Inches(0.55), Inches(3.3), Inches(0.6))
            p_p = tb_p.text_frame.paragraphs[0]
            p_p.alignment = PP_ALIGN.RIGHT
            r_p = p_p.add_run()
            r_p.text = f"🎙️ {presenter}"
            r_p.font.size = Pt(11.5)
            r_p.font.bold = True
            r_p.font.color.rgb = ACCENT_BLUE

    def create_card(slide, left, top, width, height, bg=PANEL_DARK, border=PANEL_BORDER, lw=1.2):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = border
        card.line.width = Pt(lw)
        return card

    # =========================================================================
    # SLIDE 1: BIG IMPACT OPENING (Hero Title)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, RGBColor(7, 10, 19))

    # Glow / Hero panel behind title
    hero_card = create_card(s1, Inches(0.8), Inches(0.8), Inches(11.733), Inches(4.3), 
                            bg=RGBColor(15, 23, 42), border=RGBColor(30, 41, 59), lw=2)

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.1), Inches(11.0), Inches(3.8))
    tf1 = tb1.text_frame

    p_badge = tf1.paragraphs[0]
    r_badge = p_badge.add_run()
    r_badge.text = "OPERATING SYSTEMS PBL SHOWCASE • 20% IA WEIGHTAGE"
    r_badge.font.size = Pt(12)
    r_badge.font.bold = True
    r_badge.font.color.rgb = ACCENT_ORANGE
    p_badge.space_after = Pt(12)

    p_main = tf1.add_paragraph()
    r_main = p_main.add_run()
    r_main.text = "CPU Scheduling Beyond Textbooks"
    r_main.font.size = Pt(36)
    r_main.font.bold = True
    r_main.font.color.rgb = WHITE
    p_main.space_after = Pt(4)

    p_sub = tf1.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Modeling Real Kernel Mechanics: Multi-Burst Cycles, Dispatcher Latency (δ), and Real-Time Bounds"
    r_sub.font.size = Pt(17)
    r_sub.font.color.rgb = ACCENT_BLUE
    p_sub.space_after = Pt(20)

    # 3 Feature Pills
    p_pills = tf1.add_paragraph()
    pills = [
        (" [1] CPU1 → I/O → CPU2 ", ACCENT_BLUE),
        (" [2] Context Switch Cost (δ) ", ACCENT_RED),
        (" [3] Real-Time Bounds (RM + EDF) ", ACCENT_GREEN)
    ]
    for text, col in pills:
        r_pil = p_pills.add_run()
        r_pil.text = text
        r_pil.font.bold = True
        r_pil.font.size = Pt(13)
        r_pil.font.color.rgb = col
        r_sp = p_pills.add_run()
        r_sp.text = "   "

    # Bottom Team Roster Card
    team_card = create_card(s1, Inches(0.8), Inches(5.3), Inches(11.733), Inches(1.7), 
                            bg=PANEL_DARK, border=RGBColor(38, 50, 72))
    tb_team = s1.shapes.add_textbox(Inches(1.0), Inches(5.38), Inches(11.3), Inches(1.5))
    tf_team = tb_team.text_frame

    p_th = tf_team.paragraphs[0]
    r_th = p_th.add_run()
    r_th.text = "👥 Project Team (5 Members) | B.Tech CSE:"
    r_th.font.size = Pt(12)
    r_th.font.bold = True
    r_th.font.color.rgb = LIGHT_GRAY
    p_th.space_after = Pt(4)

    team_members = [
        "• Sparsh Srivastava (590022251) - Lead & Theoretical Architect",
        "• Aayushman Kothari (590025499) - Preemptive Schedulers & Aging",
        "• Yasharth Jaiswal (590023687) - Round Robin & Context Switch Latency (δ)",
        "• Satvik Arora (590028424) - Real-Time Systems (RM & EDF Bounds)",
        "• Yashvardhan Singh (590021777) - Architecture, Web UI & Benchmarking"
    ]
    for m in team_members:
        pm = tf_team.add_paragraph()
        rm = pm.add_run()
        rm.text = m
        rm.font.size = Pt(10)
        rm.font.color.rgb = GRAY_TEXT

    # =========================================================================
    # SLIDE 2: WHY EXISTING SIMULATORS ARE WRONG (Split Comparison)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_slide_header(s2, "Why Traditional Scheduling Simulators Fail", presenter="Sparsh Srivastava")

    # Left: Textbook Flaws Card
    create_card(s2, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3), 
                bg=RGBColor(24, 18, 26), border=RGBColor(239, 68, 68), lw=1.8)
    tb_left = s2.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.2), Inches(4.8))
    tf_l = tb_left.text_frame
    p_lh = tf_l.paragraphs[0]
    r_lh = p_lh.add_run()
    r_lh.text = "❌ Standard Textbook Assumptions"
    r_lh.font.size = Pt(17)
    r_lh.font.bold = True
    r_lh.font.color.rgb = ACCENT_RED
    p_lh.space_after = Pt(14)

    flaws = [
        ("Single Monolithic Burst:", "Processes modeled as single continuous CPU runs. Completely ignores peripheral I/O waits and device wait queues."),
        ("Zero Context Switch Cost (δ = 0):", "Assumes CPU switches processes instantaneously with 0 ms dispatcher overhead. Hides real CPU thrashing."),
        ("No Starvation Resolution:", "Priority queues are static; low-priority tasks wait indefinitely with zero mitigation."),
        ("No Real-Time Guarantees:", "Arbitrary priority schemes with zero verification against hard mathematical deadline criteria.")
    ]
    for title, desc in flaws:
        p = tf_l.add_paragraph()
        r_b = p.add_run()
        r_b.text = f"• {title}\n  "
        r_b.font.bold = True
        r_b.font.size = Pt(11)
        r_b.font.color.rgb = WHITE
        r_d = p.add_run()
        r_d.text = desc
        r_d.font.size = Pt(10)
        r_d.font.color.rgb = GRAY_TEXT
        p.space_after = Pt(8)

    # Right: Our Kernel Reality Card
    create_card(s2, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3), 
                bg=RGBColor(16, 28, 28), border=RGBColor(16, 185, 129), lw=1.8)
    tb_right = s2.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(4.8))
    tf_r = tb_right.text_frame
    p_rh = tf_r.paragraphs[0]
    r_rh = p_rh.add_run()
    r_rh.text = "✅ Our Production Kernel Engine"
    r_rh.font.size = Pt(17)
    r_rh.font.bold = True
    r_rh.font.color.rgb = ACCENT_GREEN
    p_rh.space_after = Pt(14)

    solutions = [
        ("Multi-Burst Interleaved Model:", "Models CPU1 → I/O Trapping → Device Queue → Interrupt → CPU2. Accurately simulates the 5-state lifecycle."),
        ("Explicit Context Switch Latency (δ):", "Injects discrete δ latency per preemption to save/restore PCB registers and quantify CPU Efficiency (η)."),
        ("Dynamic Priority Aging Engine:", "Automatically boosts waiting processes each epoch, mathematically guaranteeing bounded waiting."),
        ("Liu-Layland Real-Time Verification:", "Automated schedulability evaluation checking Rate Monotonic (RM) and Earliest Deadline First (EDF) bounds.")
    ]
    for title, desc in solutions:
        p = tf_r.add_paragraph()
        r_b = p.add_run()
        r_b.text = f"• {title}\n  "
        r_b.font.bold = True
        r_b.font.color.rgb = WHITE
        r_d = p.add_run()
        r_d.text = desc
        r_d.font.size = Pt(10)
        r_d.font.color.rgb = GRAY_TEXT
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 3: PROCESS LIFECYCLE ENGINE (Visual State Diagram)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_slide_header(s3, "Process Lifecycle Engine: 5-State Multi-Burst Machine", presenter="Sparsh Srivastava")

    # State Box Helper
    def draw_state_node(slide, left, top, w, h, name, desc, color):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(17, 24, 39)
        card.line.color.rgb = color
        card.line.width = Pt(2)

        tb = slide.shapes.add_textbox(left, top, w, h)
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = name
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = color
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = GRAY_TEXT

    # 1. NEW
    draw_state_node(s3, Inches(0.8), Inches(3.2), Inches(1.8), Inches(1.2), "NEW", "Program Admitted\n(Long-Term)", ACCENT_BLUE)
    # 2. READY
    draw_state_node(s3, Inches(3.6), Inches(3.2), Inches(2.2), Inches(1.2), "READY", "RAM Ready Queue\n(Scheduling Pool)", ACCENT_GREEN)
    # 3. RUNNING
    draw_state_node(s3, Inches(7.0), Inches(3.2), Inches(2.2), Inches(1.2), "RUNNING", "CPU Allocation\n(Executing CPU1/2)", ACCENT_ORANGE)
    # 4. TERMINATED
    draw_state_node(s3, Inches(10.4), Inches(3.2), Inches(2.1), Inches(1.2), "TERMINATED", "PCB Reclaimed\n(Exit Trap)", ACCENT_PURPLE)
    # 5. WAITING (Below)
    draw_state_node(s3, Inches(5.3), Inches(5.4), Inches(2.6), Inches(1.2), "WAITING (I/O)", "Device Wait Queue\n(I/O Burst Countdown)", ACCENT_RED)

    # Flow arrows via text banners
    # New -> Ready
    tb_arr1 = s3.shapes.add_textbox(Inches(2.6), Inches(3.4), Inches(1.0), Inches(0.8))
    p = tb_arr1.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "──▶\nAdmit"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = ACCENT_BLUE

    # Ready -> Running
    tb_arr2 = s3.shapes.add_textbox(Inches(5.8), Inches(3.4), Inches(1.2), Inches(0.8))
    p = tb_arr2.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "──▶\nDispatch"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = ACCENT_GREEN

    # Running -> Terminated
    tb_arr3 = s3.shapes.add_textbox(Inches(9.2), Inches(3.4), Inches(1.2), Inches(0.8))
    p = tb_arr3.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "──▶\nFinish"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PURPLE

    # Running -> Waiting (Down)
    tb_arr4 = s3.shapes.add_textbox(Inches(7.3), Inches(4.45), Inches(1.8), Inches(0.8))
    p = tb_arr4.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "▼ Syscall Trap (I/O Req)"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = ACCENT_RED

    # Waiting -> Ready (Up)
    tb_arr5 = s3.shapes.add_textbox(Inches(3.8), Inches(4.45), Inches(1.8), Inches(0.8))
    p = tb_arr5.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "▲ I/O Done (Interrupt)"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = ACCENT_GREEN

    # Top Explanation Box
    box_info = create_card(s3, Inches(0.8), Inches(1.5), Inches(11.733), Inches(1.3),
                           bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_top = s3.shapes.add_textbox(Inches(1.0), Inches(1.55), Inches(11.3), Inches(1.2))
    tf_top = tb_top.text_frame
    p_t = tf_top.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "💡 Architectural Multi-Burst Interleaving: CPU1 → I/O → CPU2"
    r_t.font.size = Pt(13)
    r_t.font.bold = True
    r_t.font.color.rgb = ACCENT_BLUE
    p_t.space_after = Pt(4)

    p_td = tf_top.add_paragraph()
    r_td = p_td.add_run()
    r_td.text = "When P1 finishes CPU1, it triggers a software trap transitioning User Mode (1) → Kernel Mode (0). P1 enters the Device Queue (WAITING), unblocking the Ready Queue for P2. Upon peripheral completion, a hardware interrupt fires an ISR, re-enqueueing P1 for CPU2."
    r_td.font.size = Pt(10.5)
    r_td.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 4: INSIDE OUR SCHEDULER (Pipeline Architecture)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_slide_header(s4, "Scheduler Engine Architecture Pipeline", presenter="Yashvardhan Singh")

    pipeline_stages = [
        ("1. Input Ingestion", "Process Table\nBurst Vectors\n(CPU1, IO, CPU2)\nPriorities & Periods", ACCENT_BLUE),
        ("2. Kernel Simulator", "Cycle-accurate clock\n5-State PCB Manager\nDevice Waiting Queue\nSoftware Trap Handler", ACCENT_GREEN),
        ("3. Policy Dispatcher", "FCFS\nSJF / SRTF\nPreemptive Priority\nRound Robin (q)", ACCENT_ORANGE),
        ("4. Context Switch Mgr", "Registers Save/Restore\nInjects δ Delay\nCalculates CPU\nEfficiency (η)", ACCENT_RED),
        ("5. Analytics & UI", "Interactive Gantt (Plotly)\nPCB Metrics Table\nLiu-Layland RM/EDF\nBenchmark Charts", ACCENT_PURPLE)
    ]

    card_w = Inches(2.15)
    card_h = Inches(3.8)
    spacing = Inches(0.24)
    start_x = Inches(0.8)
    start_y = Inches(1.8)

    for idx, (title, content, color) in enumerate(pipeline_stages):
        cur_x = start_x + idx * (card_w + spacing)
        create_card(s4, cur_x, start_y, card_w, card_h, bg=RGBColor(17, 24, 39), border=color, lw=1.8)

        tb = s4.shapes.add_textbox(cur_x + Inches(0.1), start_y + Inches(0.15), card_w - Inches(0.2), card_h - Inches(0.3))
        tf = tb.text_frame
        p_num = tf.paragraphs[0]
        r_num = p_num.add_run()
        r_num.text = f"STAGE {idx+1}"
        r_num.font.size = Pt(10)
        r_num.font.bold = True
        r_num.font.color.rgb = color
        p_num.space_after = Pt(4)

        p_h = tf.add_paragraph()
        r_h = p_h.add_run()
        r_h.text = title
        r_h.font.size = Pt(13)
        r_h.font.bold = True
        r_h.font.color.rgb = WHITE
        p_h.space_after = Pt(12)

        for line in content.split('\n'):
            p_l = tf.add_paragraph()
            r_l = p_l.add_run()
            r_l.text = f"• {line}"
            r_l.font.size = Pt(9.5)
            r_l.font.color.rgb = LIGHT_GRAY
            p_l.space_after = Pt(4)

        # Connector arrow
        if idx < len(pipeline_stages) - 1:
            arr_x = cur_x + card_w
            arr_tb = s4.shapes.add_textbox(arr_x, start_y + Inches(1.6), spacing, Inches(0.6))
            p_a = arr_tb.text_frame.paragraphs[0]
            p_a.alignment = PP_ALIGN.CENTER
            r_a = p_a.add_run()
            r_a.text = "▶"
            r_a.font.size = Pt(14)
            r_a.font.bold = True
            r_a.font.color.rgb = ACCENT_BLUE

    # Bottom Callout
    bot_card = create_card(s4, Inches(0.8), Inches(5.9), Inches(11.733), Inches(1.1),
                           bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_bot = s4.shapes.add_textbox(Inches(1.0), Inches(5.95), Inches(11.3), Inches(1.0))
    tf_b = tb_bot.text_frame
    pb = tf_b.paragraphs[0]
    rb = pb.add_run()
    rb.text = "🎯 Engineering Modularity: Discrete Separation of Kernel Engine (`engine.py`) and Visualizer (`app.py`)"
    rb.font.size = Pt(12)
    rb.font.bold = True
    rb.font.color.rgb = ACCENT_BLUE
    pb.space_after = Pt(2)

    pb2 = tf_b.add_paragraph()
    rb2 = pb2.add_run()
    rb2.text = "The core engine runs a cycle-accurate tick loop recording granular Gantt events without UI coupling. Results stream into Plotly and Streamlit for live rendering, state inspection, and cross-policy comparisons."
    rb2.font.size = Pt(10)
    rb2.font.color.rgb = GRAY_TEXT

    # =========================================================================
    # SLIDE 5: CONTEXT SWITCHING COSTS MATTER (The Big Visual Number)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_slide_header(s5, "The Cost of Preemption: Context Switch Latency (δ)", presenter="Yasharth Jaiswal")

    # Left: RR(q=2) vs Right: RR(q=4)
    # Left Box
    create_card(s5, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3),
                bg=RGBColor(24, 18, 26), border=RGBColor(239, 68, 68), lw=2)
    tb_q2 = s5.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.2), Inches(4.9))
    tf_q2 = tb_q2.text_frame

    p = tf_q2.paragraphs[0]
    r = p.add_run()
    r.text = "Round Robin with Small Quantum (q = 2, δ = 1)"
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = ACCENT_RED
    p.space_after = Pt(10)

    p_g1 = tf_q2.add_paragraph()
    r_g1 = p_g1.add_run()
    r_g1.text = "Timeline Pattern:\n| P1 | δ | P2 | δ | P3 | δ | P1 | δ | P2 |"
    r_g1.font.size = Pt(11)
    r_g1.font.bold = True
    r_g1.font.color.rgb = WHITE
    p_g1.space_after = Pt(18)

    p_num1 = tf_q2.add_paragraph()
    r_num1 = p_num1.add_run()
    r_num1.text = "η = 73.08%"
    r_num1.font.size = Pt(40)
    r_num1.font.bold = True
    r_num1.font.color.rgb = ACCENT_RED
    p_num1.space_after = Pt(6)

    p_d1 = tf_q2.add_paragraph()
    r_d1 = p_d1.add_run()
    r_d1.text = "⚠️ Severe Context Switch Thrashing!\nNearly 27% of all CPU execution cycles are lost to dispatcher register saving/restoring into the PCB rather than productive user code."
    r_d1.font.size = Pt(11)
    r_d1.font.color.rgb = LIGHT_GRAY

    # Right Box
    create_card(s5, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3),
                bg=RGBColor(16, 28, 28), border=RGBColor(16, 185, 129), lw=2)
    tb_q4 = s5.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(4.9))
    tf_q4 = tb_q4.text_frame

    p = tf_q4.paragraphs[0]
    r = p.add_run()
    r.text = "Round Robin with Tuned Quantum (q = 4, δ = 1)"
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = ACCENT_GREEN
    p.space_after = Pt(10)

    p_g2 = tf_q4.add_paragraph()
    r_g2 = p_g2.add_run()
    r_g2.text = "Timeline Pattern:\n| P1 ─────── | δ | P2 ─────── | δ |"
    r_g2.font.size = Pt(11)
    r_g2.font.bold = True
    r_g2.font.color.rgb = WHITE
    p_g2.space_after = Pt(18)

    p_num2 = tf_q4.add_paragraph()
    r_num2 = p_num2.add_run()
    r_num2.text = "η = 81.82%"
    r_num2.font.size = Pt(40)
    r_num2.font.bold = True
    r_num2.font.color.rgb = ACCENT_GREEN
    p_num2.space_after = Pt(6)

    p_d2 = tf_q4.add_paragraph()
    r_d2 = p_d2.add_run()
    r_d2.text = "✅ Optimal Efficiency Restored!\nIncreasing the quantum size reduces dispatcher frequency by 50%, recovering +8.74% of productive CPU utilization while preserving interactive responsiveness."
    r_d2.font.size = Pt(11)
    r_d2.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 6: STARVATION PROBLEM & AGING (Storytelling Layout)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_slide_header(s6, "Starvation Hazard & Dynamic Aging Solution", presenter="Aayushman Kothari")

    # Left: The Hazard Story
    create_card(s6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3),
                bg=RGBColor(24, 18, 26), border=RGBColor(239, 68, 68), lw=1.8)
    tb_s1 = s6.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.2), Inches(4.8))
    tf_s1 = tb_s1.text_frame

    p = tf_s1.paragraphs[0]
    r = p.add_run()
    r.text = "⚠️ The Indefinite Blocking Hazard"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = ACCENT_RED
    p.space_after = Pt(12)

    hazard_steps = [
        "Incoming stream of High-Priority tasks (Priority = 1)",
        "Incoming stream of High-Priority tasks (Priority = 1)",
        "Incoming stream of High-Priority tasks (Priority = 1)",
        "Low-Priority Process P5 (Priority = 5) sits in Ready Queue...",
        "⏳ WAITING FOREVER... (Starvation)"
    ]
    for step in hazard_steps:
        p_step = tf_s1.add_paragraph()
        r_s = p_step.add_run()
        r_s.text = f"• {step}"
        r_s.font.size = Pt(11)
        r_s.font.color.rgb = LIGHT_GRAY if "WAITING" not in step else ACCENT_RED
        r_s.font.bold = ("WAITING" in step)
        p_step.space_after = Pt(8)

    p_note = tf_s1.add_paragraph()
    r_note = p_note.add_run()
    r_note.text = "\nDifference from Deadlock: In starvation, resources are actively moving and being utilized by other tasks; one process is simply indefinitely bypassed."
    r_note.font.size = Pt(10)
    r_note.font.color.rgb = GRAY_TEXT

    # Right: Dynamic Aging Engine
    create_card(s6, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3),
                bg=RGBColor(16, 28, 28), border=RGBColor(16, 185, 129), lw=1.8)
    tb_s2 = s6.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(4.8))
    tf_s2 = tb_s2.text_frame

    p = tf_s2.paragraphs[0]
    r = p.add_run()
    r.text = "✅ The Dynamic Aging Engine"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = ACCENT_GREEN
    p.space_after = Pt(12)

    p_form = tf_s2.add_paragraph()
    r_form = p_form.add_run()
    r_form.text = "Mathematical Priority Escalation:\nPriority_eff = max(0, Priority_init - ⌊Wait_Time / Epoch⌋)"
    r_form.font.size = Pt(11.5)
    r_form.font.bold = True
    r_form.font.color.rgb = ACCENT_BLUE
    p_form.space_after = Pt(14)

    aging_steps = [
        "Epoch t=4: P5 priority boosted (5 ➔ 4)",
        "Epoch t=8: P5 priority boosted (4 ➔ 3)",
        "Epoch t=12: P5 priority boosted (3 ➔ 2)",
        "Epoch t=16: P5 priority reaches 1 (Highest!)",
        "🚀 P5 DISPATCHED TO CPU & COMPLETES!"
    ]
    for step in aging_steps:
        p_step = tf_s2.add_paragraph()
        r_s = p_step.add_run()
        r_s.text = f"• {step}"
        r_s.font.size = Pt(11)
        r_s.font.color.rgb = ACCENT_GREEN if "DISPATCHED" in step else LIGHT_GRAY
        r_s.font.bold = ("DISPATCHED" in step)
        p_step.space_after = Pt(8)

    # =========================================================================
    # SLIDE 7: REAL-TIME SCHEDULING (Two Giant Cards: RM & EDF)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_slide_header(s7, "Real-Time Schedulability: Rate Monotonic vs. EDF", presenter="Satvik Arora")

    # Left Card: Rate Monotonic (RM)
    create_card(s7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(4.2),
                bg=RGBColor(17, 24, 39), border=ACCENT_BLUE, lw=2)
    tb_rm = s7.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.2), Inches(3.8))
    tf_rm = tb_rm.text_frame

    p = tf_rm.paragraphs[0]
    r = p.add_run()
    r.text = "Rate Monotonic (RM) Static Priority"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = ACCENT_BLUE
    p.space_after = Pt(4)

    p_sub = tf_rm.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Liu-Layland Utilization Bound (Static Optimal)"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = GRAY_TEXT
    p_sub.space_after = Pt(14)

    p_num = tf_rm.add_paragraph()
    r_num = p_num.add_run()
    r_num.text = "69.3%"
    r_num.font.size = Pt(46)
    r_num.font.bold = True
    r_num.font.color.rgb = ACCENT_BLUE
    p_num.space_after = Pt(8)

    p_d = tf_rm.add_paragraph()
    r_d = p_d.add_run()
    r_d.text = "Formula: U ≤ N(2^(1/N) - 1)\nConverges to ln(2) ≈ 69.3% as N ➔ ∞. If total CPU utilization U ≤ 69.3%, all task deadlines are 100% mathematically guaranteed under static priority."
    r_d.font.size = Pt(10.5)
    r_d.font.color.rgb = LIGHT_GRAY

    # Right Card: Earliest Deadline First (EDF)
    create_card(s7, Inches(6.8), Inches(1.6), Inches(5.7), Inches(4.2),
                bg=RGBColor(17, 24, 39), border=ACCENT_GREEN, lw=2)
    tb_edf = s7.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(3.8))
    tf_edf = tb_edf.text_frame

    p = tf_edf.paragraphs[0]
    r = p.add_run()
    r.text = "Earliest Deadline First (EDF) Dynamic Priority"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = ACCENT_GREEN
    p.space_after = Pt(4)

    p_sub = tf_edf.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Dynamic Optimal Theoretical Limit"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = GRAY_TEXT
    p_sub.space_after = Pt(14)

    p_num = tf_edf.add_paragraph()
    r_num = p_num.add_run()
    r_num.text = "100.0%"
    r_num.font.size = Pt(46)
    r_num.font.bold = True
    r_num.font.color.rgb = ACCENT_GREEN
    p_num.space_after = Pt(8)

    p_d = tf_edf.add_paragraph()
    r_d = p_d.add_run()
    r_d.text = "Formula: U = Σ(C_i / T_i) ≤ 1.00\nEDF is dynamically optimal. By prioritizing jobs with nearest absolute deadline, EDF can safely utilize 100% of CPU cycles without missing a single deadline."
    r_d.font.size = Pt(10.5)
    r_d.font.color.rgb = LIGHT_GRAY

    # Bottom Combined Verdict Card
    bot_card = create_card(s7, Inches(0.8), Inches(6.0), Inches(11.733), Inches(1.1),
                           bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_ver = s7.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.3), Inches(1.0))
    tf_v = tb_ver.text_frame
    pv = tf_v.paragraphs[0]
    rv = pv.add_run()
    rv.text = "🔬 Automated Test Task Set Verification in Our Simulator:"
    rv.font.bold = True
    rv.font.size = Pt(11.5)
    rv.font.color.rgb = ACCENT_ORANGE
    pv.space_after = Pt(2)

    pv2 = tf_v.add_paragraph()
    rv2 = pv2.add_run()
    rv2.text = "Test workload (P1–P5) generated total utilization U = 162.08%. Since U > 100%, our engine instantly fired an automated Infeasibility Alert: both RM and EDF prove that the task set is over-utilized and deadlines will be missed."
    rv2.font.size = Pt(10)
    rv2.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 8: LIVE SIMULATOR DEMONSTRATION (PROOF & SCREENSHOT)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_slide_header(s8, "Live Kernel Simulator Showcase (Engineering Proof)", presenter="Yashvardhan Singh")

    sim_img_path = os.path.join(assets_dir, "simulator_screenshot.png")
    if os.path.exists(sim_img_path):
        s8.shapes.add_picture(sim_img_path, Inches(0.8), Inches(1.5), width=Inches(11.733), height=Inches(5.4))
    else:
        # Fallback card
        create_card(s8, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.4))

    # =========================================================================
    # SLIDE 9: BENCHMARK RESULTS (Data Charts - Not Boring Table)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_slide_header(s9, "Empirical Benchmarks: Latency vs. CPU Efficiency", presenter="Aayushman & Yasharth")

    bench_img_path = os.path.join(assets_dir, "benchmark_charts.png")
    if os.path.exists(bench_img_path):
        s9.shapes.add_picture(bench_img_path, Inches(0.8), Inches(1.5), width=Inches(11.733), height=Inches(4.5))

    # Bottom Summary Takeaways
    b_takeaway = create_card(s9, Inches(0.8), Inches(6.1), Inches(11.733), Inches(1.0),
                             bg=RGBColor(15, 23, 42), border=RGBColor(38, 50, 72))
    tb_t = s9.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(11.3), Inches(0.9))
    tf_t = tb_t.text_frame
    pt = tf_t.paragraphs[0]
    rt = pt.add_run()
    rt.text = "📊 Key Takeaways: "
    rt.font.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = ACCENT_BLUE

    rt_desc = pt.add_run()
    rt_desc.text = "1. Priority with Aging achieved the lowest response latency (1.80 units).  2. Round Robin at q=2 incurred heavy preemption costs (73.08% efficiency).  3. Increasing quantum to q=4 restored efficiency to 81.82% without degrading responsiveness."
    rt_desc.font.size = Pt(10)
    rt_desc.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 10: WHY OUR PROJECT STANDS OUT (4 Hero Pillars & Mic Drop)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_slide_header(s10, "Why Our Project Stands Out: Faculty Evaluation Matrix", presenter="All 5 Team Members")

    card_pillars = [
        ("✅ 1. Multi-Burst Engine", "Models real CPU1 ➔ I/O ➔ CPU2 transitions. Accurately tracks device wait queues and interrupt re-admission.", ACCENT_BLUE),
        ("✅ 2. Context Switch (δ)", "Explicitly quantifies dispatcher register save/restore overhead on PCB; uncovers CPU thrashing.", ACCENT_RED),
        ("✅ 3. Dynamic Aging", "Mathematically proves starvation prevention by dynamically elevating process priority over time.", ACCENT_GREEN),
        ("✅ 4. Real-Time Bounds", "Implements Liu-Layland mathematical verification (RM ≤ 69.3% vs. EDF ≤ 100%) for periodic deadlines.", ACCENT_PURPLE)
    ]

    p_w = Inches(2.72)
    p_h = Inches(3.2)
    p_gap = Inches(0.28)
    p_x = Inches(0.8)
    p_y = Inches(1.6)

    for i, (title, text, color) in enumerate(card_pillars):
        c_x = p_x + i * (p_w + p_gap)
        create_card(s10, c_x, p_y, p_w, p_h, bg=RGBColor(17, 24, 39), border=color, lw=2)

        tb = s10.shapes.add_textbox(c_x + Inches(0.12), p_y + Inches(0.15), p_w - Inches(0.24), p_h - Inches(0.3))
        tf = tb.text_frame
        p_c = tf.paragraphs[0]
        r_c = p_c.add_run()
        r_c.text = title
        r_c.font.size = Pt(13)
        r_c.font.bold = True
        r_c.font.color.rgb = color
        p_c.space_after = Pt(10)

        p_desc = tf.add_paragraph()
        r_desc = p_desc.add_run()
        r_desc.text = text
        r_desc.font.size = Pt(10)
        r_desc.font.color.rgb = LIGHT_GRAY

    # Big Mic Drop Banner
    drop_card = create_card(s10, Inches(0.8), Inches(5.1), Inches(11.733), Inches(1.9),
                            bg=RGBColor(15, 23, 42), border=ACCENT_ORANGE, lw=2)
    tb_drop = s10.shapes.add_textbox(Inches(1.0), Inches(5.2), Inches(11.3), Inches(1.7))
    tf_drop = tb_drop.text_frame

    p_drop = tf_drop.paragraphs[0]
    p_drop.alignment = PP_ALIGN.CENTER
    r_drop = p_drop.add_run()
    r_drop.text = "🎯 Closer to Real Operating System Kernels Than Traditional Simulators"
    r_drop.font.size = Pt(18)
    r_drop.font.bold = True
    r_drop.font.color.rgb = ACCENT_ORANGE
    p_drop.space_after = Pt(4)

    p_cred = tf_drop.add_paragraph()
    p_cred.alignment = PP_ALIGN.CENTER
    r_cred = p_cred.add_run()
    r_cred.text = "Demonstrated full-stack engineering: Mathematics ➔ Kernel Engine (`engine.py`) ➔ Interactive GUI (`app.py`) ➔ Empirical Benchmarking\nPrepared by: Sparsh Srivastava | Aayushman Kothari | Yasharth Jaiswal | Satvik Arora | Yashvardhan Singh"
    r_cred.font.size = Pt(10.5)
    r_cred.font.color.rgb = WHITE
    p_cred.space_after = Pt(4)

    p_close = tf_drop.add_paragraph()
    p_close.alignment = PP_ALIGN.CENTER
    r_close = p_close.add_run()
    r_close.text = "✨ Thank You. We are ready for Viva Voce examination! ✨"
    r_close.font.size = Pt(12)
    r_close.font.bold = True
    r_close.font.color.rgb = ACCENT_GREEN

    # Save
    out_pptx = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\OS_PBL_Presentation_Deck.pptx"
    prs.save(out_pptx)
    print("Top-1 Deck generated successfully:", out_pptx)

if __name__ == "__main__":
    build_top1_deck()
