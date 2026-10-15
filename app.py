import streamlit as st
import pandas as pd
import plotly.express as px
from engine import Process, OSKernelSimulator

st.set_page_config(page_title="OS PBL: CPU Scheduling Suite", layout="wide")

st.title("⚡ Operating Systems PBL: CPU Scheduling & Process Management Suite")
st.markdown("""
**Course:** Operating Systems | **PBL Assessment (20% Weightage)**  
**Team Members:**  
• **Sparsh Srivastava** (590022251) — *Theoretical Architect & Non-Preemptive Lead*  
• **Aayushman Kothari** (590025499) — *Preemptive Scheduling & Starvation Specialist*  
• **Yasharth Jaiswal** (590023687) — *Time-Sharing & Dispatcher Overhead Engineer*  
• **Satvik Arora** (590028424) — *Real-Time Systems & Mathematical Proof Lead*  
• **Yashvardhan Singh** (590021777) — *System Architecture, GUI & Benchmarking Lead*  
""")
st.caption("Covering Units I & II: Multi-Burst (CPU1 → I/O → CPU2), Context Switch Overhead (δ), Real-Time Feasibility (RM & EDF)")

st.sidebar.header("🛠️ Kernel Parameters")
cs_delta = st.sidebar.slider("Context Switch Overhead (δ)", min_value=0, max_value=3, value=1, help="Simulates dispatcher save/restore latency.")
rr_quantum = st.sidebar.slider("Round Robin Time Quantum (q)", min_value=1, max_value=8, value=2)
enable_aging = st.sidebar.checkbox("Enable Dynamic Aging (Priority)", value=True)
aging_interval = st.sidebar.number_input("Aging Interval (time units)", min_value=1, value=4)

default_data = pd.DataFrame([
    {"PID": "P1", "Arrival": 0, "CPU_Burst_1": 4, "IO_Burst": 2, "CPU_Burst_2": 2, "Priority": 2, "Period": 12},
    {"PID": "P2", "Arrival": 1, "CPU_Burst_1": 3, "IO_Burst": 1, "CPU_Burst_2": 1, "Priority": 1, "Period": 8},
    {"PID": "P3", "Arrival": 2, "CPU_Burst_1": 2, "IO_Burst": 0, "CPU_Burst_2": 0, "Priority": 4, "Period": 6},
    {"PID": "P4", "Arrival": 3, "CPU_Burst_1": 1, "IO_Burst": 0, "CPU_Burst_2": 0, "Priority": 3, "Period": 10},
    {"PID": "P5", "Arrival": 4, "CPU_Burst_1": 2, "IO_Burst": 1, "CPU_Burst_2": 1, "Priority": 5, "Period": 16},
])

st.subheader("📝 Process Table & PCB Configuration")
st.info("💡 Supports Multi-Burst execution (CPU1 → I/O → CPU2) as specified in Unit II Classroom Special Variations.")
edited_df = st.data_editor(default_data, num_rows="dynamic", use_container_width=True)

procs = []
for _, r in edited_df.iterrows():
    procs.append(Process(
        pid=str(r["PID"]),
        arrival_time=int(r["Arrival"]),
        cpu_burst_1=int(r["CPU_Burst_1"]),
        io_burst=int(r["IO_Burst"]),
        cpu_burst_2=int(r["CPU_Burst_2"]),
        priority=int(r["Priority"]),
        period=int(r["Period"])
    ))

tab1, tab2, tab3 = st.tabs(["📊 Scheduling Simulation", "⏱️ Real-Time Feasibility (RM & EDF)", "🔬 Benchmark Comparison"])

with tab1:
    algo_choice = st.selectbox("Select Policy", [
        "First-Come, First-Served (FCFS) [Multi-Burst + I/O]",
        "Preemptive Priority (with Aging)",
        "Round Robin (RR) with Context Switch Overhead (δ)"
    ])

    if st.button("Simulate Schedule", type="primary"):
        if "FCFS" in algo_choice:
            gantt, results, eff = OSKernelSimulator.simulate_fcfs(procs, cs_delta)
        elif "Priority" in algo_choice:
            gantt, results, eff = OSKernelSimulator.simulate_priority(procs, enable_aging, aging_interval, cs_delta)
        else:
            gantt, results, eff = OSKernelSimulator.simulate_round_robin(procs, rr_quantum, cs_delta)

        n = len(results)
        avg_tat = round(sum(p.turnaround_time for p in results) / n, 2)
        avg_wt = round(sum(p.waiting_time for p in results) / n, 2)
        avg_rt = round(sum(p.response_time for p in results) / n, 2)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Avg Turnaround Time (TAT)", f"{avg_tat} units")
        c2.metric("Avg Waiting Time (WT)", f"{avg_wt} units")
        c3.metric("Avg Response Time (RT)", f"{avg_rt} units")
        c4.metric("CPU Efficiency (η)", f"{eff}%")

        st.subheader("Execution Gantt Chart")
        gantt_df = [dict(Task=g.label, Start=g.start, Finish=g.end, Type=g.event_type) for g in gantt]
        fig = px.timeline(pd.DataFrame(gantt_df), x_start="Start", x_end="Finish", y="Task", color="Task",
                          title=f"CPU Allocation Timeline (Context Switch Overhead δ = {cs_delta})")
        fig.update_yaxes(autorange="reversed")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Process Control Block (PCB) Termination State")
        res_table = pd.DataFrame([{
            "PID": p.pid,
            "Arrival": p.arrival_time,
            "Total CPU": p.total_cpu_time,
            "I/O Burst": p.io_burst,
            "Completion (CT)": p.completion_time,
            "Turnaround (TAT)": p.turnaround_time,
            "Waiting (WT)": p.waiting_time,
            "Response (RT)": p.response_time
        } for p in results]).sort_values(by="PID")
        st.dataframe(res_table, use_container_width=True)

with tab2:
    st.subheader("Real-Time Scheduling Feasibility Analysis (Unit II)")
    st.markdown("""
    Evaluates CPU utilization $U = \sum_{i=1}^N \\frac{C_i}{T_i}$ against:
    - **Rate Monotonic (RM) Static Bound:** $U \le N(2^{1/N} - 1)$ (Approaches $69.3\%$)
    - **Earliest Deadline First (EDF) Dynamic Bound:** $U \le 1.00$ ($100\%$)
    """)
    rt_res = OSKernelSimulator.real_time_feasibility_check(procs)
    if rt_res["valid"]:
        colA, colB, colC = st.columns(3)
        colA.metric("Total Utilization (U)", f"{rt_res['utilization'] * 100:.2f}%")
        colB.metric("RM Liu-Layland Bound", f"{rt_res['rm_bound'] * 100:.2f}%")
        colC.metric("EDF Bound", f"{rt_res['edf_bound'] * 100:.2f}%")

        if rt_res["rm_schedulable"]:
            st.success("✅ **Guaranteed Schedulable under Rate Monotonic (RM)**: $U \le N(2^{1/N}-1)$ satisfied.")
        else:
            st.warning("⚠️ **RM Infeasible or Inconclusive**: $U > N(2^{1/N}-1)$. May miss deadlines under static priorities.")

        if rt_res["edf_schedulable"]:
            st.success("✅ **Guaranteed Schedulable under EDF**: $U \le 1.00$ satisfied.")
        else:
            st.error("❌ **Unfeasible under EDF**: CPU over-utilized ($U > 1.00$). Deadlines will be missed.")

with tab3:
    st.subheader("Comprehensive Policy Benchmarking")
    if st.button("Run Multi-Policy Comparison", type="secondary"):
        _, r_fcfs, eff_fcfs = OSKernelSimulator.simulate_fcfs(procs, cs_delta)
        _, r_prio, eff_prio = OSKernelSimulator.simulate_priority(procs, enable_aging, aging_interval, cs_delta)
        _, r_rr, eff_rr = OSKernelSimulator.simulate_round_robin(procs, rr_quantum, cs_delta)

        comparison = [
            {"Policy": "FCFS (Multi-Burst)", "Avg TAT": round(sum(p.turnaround_time for p in r_fcfs)/len(procs), 2),
             "Avg WT": round(sum(p.waiting_time for p in r_fcfs)/len(procs), 2), "Efficiency (%)": eff_fcfs},
            {"Policy": "Preemptive Priority (Aging)", "Avg TAT": round(sum(p.turnaround_time for p in r_prio)/len(procs), 2),
             "Avg WT": round(sum(p.waiting_time for p in r_prio)/len(procs), 2), "Efficiency (%)": eff_prio},
            {"Policy": f"Round Robin (q={rr_quantum})", "Avg TAT": round(sum(p.turnaround_time for p in r_rr)/len(procs), 2),
             "Avg WT": round(sum(p.waiting_time for p in r_rr)/len(procs), 2), "Efficiency (%)": eff_rr},
        ]
        comp_df = pd.DataFrame(comparison)
        st.dataframe(comp_df, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.bar(comp_df, x="Policy", y=["Avg TAT", "Avg WT"], barmode="group", title="Latency Comparison (Lower is Better)")
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = px.bar(comp_df, x="Policy", y="Efficiency (%)", color="Policy", title="CPU Utilization Efficiency (Higher is Better)")
            st.plotly_chart(fig2, use_container_width=True)
