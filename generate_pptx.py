import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 Widescreen
    prs.slide_height = Inches(7.5)

    # Color Palette
    PRIMARY = RGBColor(24, 76, 120)     # #184C78 Deep Blue
    ACCENT = RGBColor(230, 81, 0)       # #E65100 Vibrant Orange
    DARK_TEXT = RGBColor(40, 40, 40)    # Charcoal
    LIGHT_BG = RGBColor(245, 247, 250)  # Off-white/slate
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(100, 100, 100)

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="OPERATING SYSTEMS PBL | 20% IA WEIGHTAGE", presenter_text=""):
        # Category header
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
        p_cat = cat_box.text_frame.paragraphs[0]
        r_cat = p_cat.add_run()
        r_cat.text = category_text.upper()
        r_cat.font.size = Pt(11)
        r_cat.font.bold = True
        r_cat.font.color.rgb = ACCENT

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(9.5), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        r_title = p_title.add_run()
        r_title.text = title_text
        r_title.font.size = Pt(22)
        r_title.font.bold = True
        r_title.font.color.rgb = PRIMARY

        # Presenter Tag
        if presenter_text:
            pres_box = slide.shapes.add_textbox(Inches(9.5), Inches(0.6), Inches(3.0), Inches(0.6))
            p_pres = pres_box.text_frame.paragraphs[0]
            p_pres.alignment = PP_ALIGN.RIGHT
            r_pres = p_pres.add_run()
            r_pres.text = f"🎙️ Presenter: {presenter_text}"
            r_pres.font.size = Pt(12)
            r_pres.font.bold = True
            r_pres.font.color.rgb = PRIMARY

    def add_card(slide, left, top, width, height, title, content_list, bg_color=WHITE, border_color=PRIMARY):
        # Card Shape
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)

        # Card Text
        tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True

        # Header in card
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = title
        r0.font.size = Pt(15)
        r0.font.bold = True
        r0.font.color.rgb = PRIMARY
        p0.space_after = Pt(8)

        for item in content_list:
            p = tf.add_paragraph()
            p.space_after = Pt(6)
            if isinstance(item, tuple):
                bold_txt, regular_txt = item
                r_b = p.add_run()
                r_b.text = bold_txt + " "
                r_b.font.bold = True
                r_b.font.size = Pt(11.5)
                r_b.font.color.rgb = DARK_TEXT
                
                r_r = p.add_run()
                r_r.text = regular_txt
                r_r.font.size = Pt(11.5)
                r_r.font.color.rgb = DARK_TEXT
            else:
                r = p.add_run()
                r.text = item
                r.font.size = Pt(11.5)
                r.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 1: TITLE SLIDE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    # Background banner
    banner = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    banner.fill.solid()
    banner.fill.fore_color.rgb = PRIMARY
    banner.line.fill.background()

    # Title box
    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(3.0))
    tf1 = tbox.text_frame
    p_sub = tf1.paragraphs[0]
    r_sub = p_sub.add_run()
    r_sub.text = "OPERATING SYSTEMS COURSE PROJECT-BASED LEARNING (PBL)"
    r_sub.font.size = Pt(14)
    r_sub.font.bold = True
    r_sub.font.color.rgb = ACCENT
    p_sub.space_after = Pt(10)

    p_t = tf1.add_paragraph()
    r_t = p_t.add_run()
    r_t.text = "Comparative Study and Simulation of CPU Scheduling Mechanics"
    r_t.font.size = Pt(32)
    r_t.font.bold = True
    r_t.font.color.rgb = WHITE
    p_t.space_after = Pt(8)

    p_st = tf1.add_paragraph()
    r_st = p_st.add_run()
    r_st.text = "Multi-Burst I/O Cycles, Context Switch Overhead (δ), and Real-Time Schedulability Analysis"
    r_st.font.size = Pt(18)
    r_st.font.italic = True
    r_st.font.color.rgb = RGBColor(200, 225, 245)

    # Team members card
    team_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.5), Inches(11.333), Inches(2.2))
    team_box.fill.solid()
    team_box.fill.fore_color.rgb = WHITE
    team_box.line.color.rgb = ACCENT
    team_box.line.width = Pt(2)

    tb_t = s1.shapes.add_textbox(Inches(1.2), Inches(4.6), Inches(11.0), Inches(2.0))
    tft = tb_t.text_frame
    pt0 = tft.paragraphs[0]
    rt0 = pt0.add_run()
    rt0.text = "👥 Project Team (Group of Exactly 5 Students) & Role Distribution:"
    rt0.font.bold = True
    rt0.font.size = Pt(13)
    rt0.font.color.rgb = PRIMARY
    pt0.space_after = Pt(6)

    members = [
        "• Sparsh Srivastava (590022251): Theoretical Architect (FCFS, SJF, Multi-Burst)",
        "• Aayushman Kothari (590025499): Preemptive Lead (SRTF, Priority, Aging Engine)",
        "• Yasharth Jaiswal (590023687): Concurrency Engineer (Round Robin, Overhead δ, Efficiency η)",
        "• Satvik Arora (590028424): Real-Time Systems Specialist (RM & EDF Liu-Layland Bounds)",
        "• Yashvardhan Singh (590021777): Architecture & UI Lead (Streamlit Engine, Plotly Gantt, PCB State)"
    ]
    for m in members:
        pm = tft.add_paragraph()
        rm = pm.add_run()
        rm.text = m
        rm.font.size = Pt(11)
        rm.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 2: MOTIVATION & PROBLEM STATEMENT
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Problem Statement: Bridging Theory vs. Kernel Reality", presenter_text="Sparsh Srivastava")

    add_card(s2, Inches(0.8), Inches(1.6), Inches(3.6), Inches(5.2), "1. The Multi-Burst Reality", [
        ("Textbook Oversimplification:", "Processes are modeled as monolithic, single CPU execution bursts."),
        ("Kernel Reality:", "Processes cycle dynamically: CPU1 -> Peripheral I/O Request -> CPU2."),
        ("Impact:", "Failing to model I/O wait queues distorts realistic ready queue arrival patterns.")
    ])

    add_card(s2, Inches(4.8), Inches(1.6), Inches(3.6), Inches(5.2), "2. Context Switch Latency (δ)", [
        ("Textbook Oversimplification:", "Dispatcher latency and state save/restore are assumed to be 0 ms."),
        ("Kernel Reality:", "Every preemption requires saving CPU registers to PCB and loading new task states."),
        ("Impact:", "As Round Robin quantum q approaches δ, context switch thrashing degrades CPU efficiency.")
    ])

    add_card(s2, Inches(8.8), Inches(1.6), Inches(3.6), Inches(5.2), "3. Real-Time Feasibility", [
        ("Textbook Oversimplification:", "Priority is arbitrary; deadlines are neglected."),
        ("Kernel Reality:", "Embedded and real-time systems require mathematical schedulability guarantees."),
        ("Impact:", "Without Rate Monotonic and EDF bounds, systems experience catastrophic deadline misses.")
    ])

    # ==========================================
    # SLIDE 3: PROCESS LIFECYCLE & PCB ANATOMY
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Unit I & II Foundations: Memory Anatomy & 5-State PCB", presenter_text="Yashvardhan Singh")

    add_card(s3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Memory Anatomy & Protection (Unit I)", [
        ("Text Segment:", "Read-only compiled machine code instructions."),
        ("Data Segment:", "Global and static initialized/uninitialized variables."),
        ("Heap Segment:", "Dynamically allocated memory (malloc/new) growing upwards."),
        ("Stack Segment:", "Local variables, activation records, growing downwards."),
        ("Dual-Mode Operation:", "User Mode (Bit = 1) vs. Kernel Mode (Bit = 0). Privileged instructions restricted. System calls execute software traps.")
    ])

    add_card(s3, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "Process Control Block & 5-State Model", [
        ("PCB Fields:", "PID, Program Counter (PC), CPU Registers, State, Priority."),
        ("State 1 - NEW:", "Program being admitted by the Long-Term Scheduler."),
        ("State 2 - READY:", "Resides in RAM, awaiting Short-Term CPU dispatch."),
        ("State 3 - RUNNING:", "Actively executing machine code on the processor."),
        ("State 4 - WAITING:", "Suspended in Device Waiting Queue awaiting I/O completion."),
        ("State 5 - TERMINATED:", "Execution concluded; resources deallocated.")
    ])

    # ==========================================
    # SLIDE 4: MULTI-BURST EXECUTION MECHANICS
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Multi-Burst Mechanics (CPU1 → I/O → CPU2)", presenter_text="Sparsh Srivastava")

    add_card(s4, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Multi-Burst Lifecycle Dynamics", [
        ("Phase 1 - Initial Execution:", "Process executes CPU_Burst_1 in RUNNING state."),
        ("Phase 2 - Trapping to I/O:", "Process invokes I/O system call, triggering software trap (Mode 1 -> 0). Process moves from RUNNING to WAITING."),
        ("Phase 3 - Device Waiting Queue:", "The process decrements I/O burst independently while other tasks utilize the CPU."),
        ("Phase 4 - Interrupt Re-admission:", "Upon I/O completion, hardware interrupt triggers ISR, moving process back to READY queue for CPU_Burst_2.")
    ])

    add_card(s4, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "Core Mathematical Metrics", [
        ("Turnaround Time (TAT):", "TAT = Completion Time (CT) - Arrival Time (AT)"),
        ("Waiting Time (WT):", "WT = TAT - Total CPU Burst - Total I/O Burst"),
        ("Response Time (RT):", "RT = First CPU Dispatch Timestamp - Arrival Time (AT)"),
        ("Convoy Effect Mitigation:", "When a long CPU-bound task traps for I/O, the ready queue is liberated, allowing short tasks to execute.")
    ])

    # ==========================================
    # SLIDE 5: PREEMPTION & AGING ENGINE
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Preemptive Schedulers & Starvation Mitigation", presenter_text="Aayushman Kothari")

    add_card(s5, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Preemptive Policies (SRTF & Priority)", [
        ("SRTF (Preemptive SJF):", "CPU is allocated to the task with shortest remaining runtime. Preempts current job on new arrivals."),
        ("Preemptive Priority:", "Lower numerical value = Higher priority. Preempts running task if a higher-priority task arrives."),
        ("The Starvation Hazard:", "Low-priority tasks can remain in the Ready Queue indefinitely if higher-priority tasks continually arrive (indefinite blocking).")
    ])

    add_card(s5, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "The Dynamic Aging Solution", [
        ("Mathematical Formulation:", "Priority_effective = Priority_initial - floor(Wait_Time / Aging_Epoch)"),
        ("Aging Mechanism:", "Every k elapsed clock ticks, waiting tasks receive a priority boost (decrement integer)."),
        ("Starvation Elimination:", "Guarantees every process eventually achieves highest priority and executes in bounded time."),
        ("Deadlock Distinction:", "Starvation is indefinite postponement (resource is moving); Deadlock is a circular hold-and-wait deadlock.")
    ])

    # ==========================================
    # SLIDE 6: ROUND ROBIN & CONTEXT SWITCH LATENCY
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Round Robin & Context Switch Penalty (δ)", presenter_text="Yasharth Jaiswal")

    add_card(s6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Time-Sharing & Quantum Dynamics", [
        ("Round Robin Policy:", "Ready queue treated as a circular FIFO. Each task is allocated a fixed time quantum q."),
        ("Extreme Case 1 (q -> ∞):", "Round Robin degenerates into First-Come First-Served (FCFS)."),
        ("Extreme Case 2 (q -> 0):", "Processor sharing abstraction; however, context switch overhead dominates physical execution."),
        ("Dispatcher Latency:", "Time required for the dispatcher to stop one process and start another.")
    ])

    add_card(s6, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "CPU Efficiency Formulation (η)", [
        ("Efficiency Formula:", "η = (Total Productive CPU Time / Total Makespan) × 100%"),
        ("Overhead Cost (δ):", "Whenever preemption occurs, δ units of CPU time are consumed saving/restoring registers."),
        ("Empirical Findings:", "With δ = 1, setting q = 2 drops CPU efficiency to 73.08%."),
        ("Optimal Tuning:", "Increasing quantum to q = 4 restores efficiency to 81.82% without significantly degrading responsiveness.")
    ])

    # ==========================================
    # SLIDE 7: REAL-TIME SCHEDULABILITY (RM & EDF)
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Real-Time Schedulability: Rate Monotonic & EDF", presenter_text="Satvik Arora")

    add_card(s7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Rate Monotonic (RM) Static Bound", [
        ("Static Assignment:", "Priorities assigned inversely to period duration (shorter period = higher priority)."),
        ("Utilization Bound:", "U = Σ(C_i / T_i) ≤ N(2^(1/N) - 1)"),
        ("Liu-Layland Proof:", "For N=1: 100%, N=2: 82.8%, N=3: 77.9%... Converges to ln(2) ≈ 69.3% as N -> ∞."),
        ("Feasibility Verdict:", "If U ≤ Bound, task set is guaranteed schedulable. If U > Bound, execution is inconclusive.")
    ])

    add_card(s7, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "Earliest Deadline First (EDF) Dynamic Bound", [
        ("Dynamic Assignment:", "Priorities assigned dynamically based on nearest absolute deadline."),
        ("Optimal Bound:", "U = Σ(C_i / T_i) ≤ 1.00 (100% Utilization)"),
        ("Feasibility Verdict:", "Task set is schedulable if and only if total utilization does not exceed 100%."),
        ("Project Automated Engine:", "Simulator calculates U against both RM and EDF bounds, issuing instant pass/fail alerts.")
    ])

    # ==========================================
    # SLIDE 8: SYSTEM ARCHITECTURE & DEMO
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Software Architecture & Interactive GUI Engine", presenter_text="Yashvardhan Singh")

    add_card(s8, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Modular Engineering Architecture", [
        ("Layer 1 - Data Layer:", "Process dataclass encapsulates PCB attributes, state, and burst vectors."),
        ("Layer 2 - Simulation Kernel:", "Discrete event-driven loop in engine.py tracking ready queues, device queues, and δ overheads."),
        ("Layer 3 - Real-Time Engine:", "Mathematical validation evaluating utilization against Liu-Layland criteria."),
        ("Layer 4 - Interactive GUI:", "Built in Streamlit and Plotly for real-time visualization.")
    ])

    add_card(s8, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "Interactive Dashboard Features", [
        ("Live Process Editor:", "Dynamic editable table supporting custom burst and real-time parameters."),
        ("Plotly Gantt Chart:", "Visualizes discrete intervals for PROCESS, CONTEXT SWITCH (δ), and IDLE states."),
        ("PCB State Inspector:", "Displays final completion, waiting, turnaround, and response metrics per task."),
        ("Side-by-Side Comparison:", "Generates comparative bar graphs across all scheduling policies.")
    ])

    # ==========================================
    # SLIDE 9: EXPERIMENTAL RESULTS & BENCHMARKING
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Experimental Findings & Comparative Benchmarks", presenter_text="Aayushman & Yasharth")

    # Table on slide
    rows, cols = 5, 5
    table_shape = s9.shapes.add_table(rows, cols, Inches(0.8), Inches(1.6), Inches(11.7), Inches(2.2))
    table = table_shape.table

    b_headers = ["Algorithm Policy", "Avg Turnaround Time", "Avg Waiting Time", "Avg Response Time", "CPU Efficiency (η)"]
    for i, h in enumerate(b_headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.bold = True
        r.font.size = Pt(12)
        r.font.color.rgb = WHITE

    b_rows = [
        ("FCFS (Multi-Burst)", "14.80 units", "8.80 units", "4.20 units", "85.71%"),
        ("Preemptive Priority (w/ Aging)", "12.40 units", "6.40 units", "1.80 units", "82.60%"),
        ("Round Robin (q=2, δ=1)", "16.20 units", "10.20 units", "2.20 units", "73.08%"),
        ("Round Robin (q=4, δ=1)", "14.60 units", "8.60 units", "3.40 units", "81.82%"),
    ]
    for r_idx, row in enumerate(b_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            bg_c = RGBColor(245, 247, 250) if r_idx % 2 == 1 else WHITE
            cell.fill.fore_color.rgb = bg_c
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.size = Pt(11.5)
            r.font.color.rgb = DARK_TEXT

    add_card(s9, Inches(0.8), Inches(4.2), Inches(11.7), Inches(2.6), "Key Empirical Takeaways", [
        ("Takeaway 1 (Quantum Tuning):", "Round Robin with q=2 incurs excessive context-switching overhead, dragging efficiency down to 73.08%. Increasing q to 4 improves efficiency by 8.74%."),
        ("Takeaway 2 (Aging Validation):", "Priority scheduling without aging resulted in starvation of P5. Dynamic aging restored fairness without degrading total makespan."),
        ("Takeaway 3 (Real-Time Infeasibility Alert):", "Task set generated U = 1.62 > 1.00; the simulator flagged infeasibility, proving deadlines cannot be met.")
    ])

    # ==========================================
    # SLIDE 10: CONCLUSION & VIVA DEFENSE
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Conclusion & Individual Contribution Summary", presenter_text="All Members")

    add_card(s10, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2), "Summary of Achievements", [
        ("Bridged Pedagogical Gap:", "Modeled multi-burst cycles, non-zero context switch overhead (δ), and real-time bounds."),
        ("Rigorous Benchmarking:", "Demonstrated trade-offs between responsiveness and CPU efficiency."),
        ("Production Quality Dashboard:", "Delivered clean Python + Streamlit + Plotly interactive platform."),
        ("Direct Syllabus Alignment:", "Covered Units I & II: Process lifecycle, dual-mode protection, PCB anatomy, and scheduling mathematics.")
    ])

    add_card(s10, Inches(6.8), Inches(1.6), Inches(5.6), Inches(5.2), "Individual Contribution & Defense Ready", [
        ("Sparsh Srivastava (590022251):", "Mathematical metrics modeling & Multi-burst FCFS/SJF lifecycle."),
        ("Aayushman Kothari (590025499):", "Preemptive SRTF, Priority scheduling & Aging starvation proof."),
        ("Yasharth Jaiswal (590023687):", "Round Robin, Context switch latency (δ) & CPU efficiency model."),
        ("Satvik Arora (590028424):", "Real-Time systems, Rate Monotonic & EDF Liu-Layland bounds."),
        ("Yashvardhan Singh (590021777):", "Full-stack Streamlit web engine, PCB inspector & benchmarking."),
        ("Open for Defense:", "The group is now ready for individual Viva Voce examination.")
    ])

    output_pptx = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\OS_PBL_Presentation_Deck.pptx"
    prs.save(output_pptx)
    print(f"Presentation saved successfully at: {output_pptx}")

if __name__ == "__main__":
    create_presentation()
