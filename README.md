# CPU Scheduling Mechanics Simulation & Benchmarking Suite
> **Course:** Operating Systems — Project-Based Learning (PBL) Component (20% IA Weightage)  
> **Topic:** Comparison and Practical Implementation of CPU Scheduling Algorithms with Multi-Burst I/O and Real-Time Feasibility.

---

## 👥 Project Team (Group of Exactly 5 Students)

| Name | Roll Number | Role |
| :--- | :---: | :--- |
| **Sparsh Srivastava** *(Lead)* | `590022251` | Theoretical Architect & Non-Preemptive Scheduling Lead |
| **Aayushman Kothari** | `590025499` | Preemptive Scheduling & Starvation Mitigation Lead |
| **Yasharth Jaiswal** | `590023687` | Concurrency & Context Switch Latency ($\delta$) Engineer |
| **Satvik Arora** | `590028424` | Real-Time Systems Specialist (RM & EDF Feasibility) |
| **Yashvardhan Singh** | `590021777` | System Architecture, UI & Benchmarking Lead |

---

## 📌 Project Overview
This project bridges the gap between theoretical textbook scheduling and real-world kernel execution by implementing an interactive simulation and benchmarking suite in **Python**, **Streamlit**, and **Plotly**.

### Key Features (Mapped to Course Syllabus Units I & II):
1. **Multi-Burst Lifecycle Simulation ($\text{CPU}_1 \to \text{I/O} \to \text{CPU}_2$):** Simulates 5-state process transitions between the Ready Queue, Running State, and Device Waiting Queue upon I/O system call traps.
2. **Context Switch Overhead ($\delta$):** Injects realistic dispatcher latency (register saving/restoring into the PCB) on every preemption to quantify CPU efficiency ($\eta$).
3. **Dynamic Priority Aging:** Prevents indefinite blocking (starvation) in priority queues by systematically boosting waiting processes.
4. **Real-Time Schedulability Verification:** Calculates CPU utilization ($U = \sum \frac{C_i}{T_i}$) against mathematical bounds:
   - **Rate Monotonic (RM):** $U \le N(2^{1/N} - 1)$ (converges to $\ln 2 \approx 69.3\%$)
   - **Earliest Deadline First (EDF):** $U \le 1.00$ ($100\%$)
5. **Interactive Web Dashboard:** Interactive Gantt timelines, live Process Control Block (PCB) inspector, and multi-algorithm comparative charts.

---

## 📂 Repository Structure

```
├── OS_PBL_Comprehensive_Report.docx   # Academic project report (Word format)
├── OS_PBL_Presentation_Deck.pptx      # 10-slide 16:9 presentation deck (PowerPoint)
├── engine.py                          # Core kernel simulation engine
├── app.py                             # Interactive Streamlit web application
├── requirements.txt                   # Python dependencies
├── generate_doc.py                    # Script used to generate the Word report
├── generate_pptx.py                   # Script used to generate the presentation deck
└── README.md                          # Project documentation
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- `pip` package manager

### Installation
1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd <repo-folder>
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the interactive web dashboard:
   ```bash
   streamlit run app.py
   ```
   The application will automatically launch at `http://localhost:8501`.

---

## 🔬 Benchmark Comparison (Sample Results)

| Algorithm Policy | Avg Turnaround Time (TAT) | Avg Waiting Time (WT) | Avg Response Time (RT) | CPU Efficiency ($\eta$) |
| :--- | :---: | :---: | :---: | :---: |
| **FCFS (Multi-Burst)** | 14.80 | 8.80 | 4.20 | 85.71% |
| **Preemptive Priority (w/ Aging)** | 12.40 | 6.40 | 1.80 | 82.60% |
| **Round Robin ($q=2, \delta=1$)** | 16.20 | 10.20 | 2.20 | 73.08% |
| **Round Robin ($q=4, \delta=1$)** | 14.60 | 8.60 | 3.40 | 81.82% |

---

## 📜 References
- A. Silberschatz, P. B. Galvin, and G. Gagne, *Operating System Concepts*, 10th ed. Wiley, 2018.
- C. L. Liu and J. W. Layland, *"Scheduling algorithms for multiprogramming in a hard-real-time environment,"* Journal of the ACM, vol. 20, no. 1, pp. 46–61, 1973.
- R. Love, *Linux Kernel Development*, 3rd ed. Addison-Wesley, 2010.
