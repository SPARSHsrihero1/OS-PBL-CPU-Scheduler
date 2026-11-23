import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

# Create high-res styling
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.family'] = 'sans-serif'

output_dir = r"C:\Users\spars\.gemini\antigravity\scratch\os_pbl_project\assets"
os.makedirs(output_dir, exist_ok=True)

# -------------------------------------------------------------
# 1. Slide 8 Screenshot: Simulator Web App Mockup (Gantt & PCB)
# -------------------------------------------------------------
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 6.2), dpi=200, gridspec_kw={'height_ratios': [1.2, 1.8]})
fig.patch.set_facecolor('#0E1117') # Streamlit Dark Mode background

# Title bar simulating Streamlit
fig.suptitle("⚡ OS Kernel CPU Scheduling Suite (Streamlit Live Dashboard)", 
             fontsize=16, fontweight='bold', color='#38BDF8', y=0.96)

# Gantt Chart Visualizer
ax1.set_facecolor('#1E293B')
ax1.set_title("Live CPU Execution Gantt Timeline (Context Switch δ = 1, Quantum q = 2)", 
              fontsize=11, fontweight='bold', color='#F8FAFC', pad=8)

# Gantt intervals: (start, duration, label, color)
segments = [
    (0, 4, "P1", "#3B82F6"),
    (4, 5, "δ (CS)", "#EF4444"),
    (5, 8, "P2", "#10B981"),
    (8, 9, "δ (CS)", "#EF4444"),
    (9, 11, "P3", "#F59E0B"),
    (11, 12, "δ (CS)", "#EF4444"),
    (12, 13, "P4", "#8B5CF6"),
    (13, 14, "δ (CS)", "#EF4444"),
    (14, 16, "P5", "#EC4899"),
    (16, 17, "δ (CS)", "#EF4444"),
    (17, 19, "P1 (CPU2)", "#3B82F6"),
    (19, 20, "δ (CS)", "#EF4444"),
    (20, 21, "P2 (CPU2)", "#10B981"),
    (21, 23, "P5 (CPU2)", "#EC4899")
]

for start, dur, label, color in segments:
    ax1.barh(0, dur, left=start, height=0.55, color=color, edgecolor='#0E1117', linewidth=1.5)
    ax1.text(start + dur/2, 0, label, ha='center', va='center', color='white', 
             fontsize=8.5, fontweight='bold')

ax1.set_xlim(-0.5, 24)
ax1.set_ylim(-0.6, 0.6)
ax1.set_yticks([])
ax1.set_xticks(range(0, 25, 2))
ax1.tick_params(colors='#94A3B8', labelsize=8.5)
ax1.set_xlabel("Timeline Clock Ticks (t)", color='#94A3B8', fontsize=9.5)
for spine in ax1.spines.values():
    spine.set_color('#334155')

# Bottom: Simulated Process Control Block & Metrics Table
ax2.set_facecolor('#1E293B')
ax2.axis('off')

# Metric Cards
metrics = [
    ("Avg Turnaround Time", "12.40 units", "#38BDF8"),
    ("Avg Waiting Time", "6.40 units", "#34D399"),
    ("Avg Response Time", "1.80 units", "#FBBF24"),
    ("CPU Efficiency (η)", "82.60%", "#A78BFA")
]

for idx, (title, val, col) in enumerate(metrics):
    x_pos = 0.05 + idx * 0.24
    rect = patches.FancyBboxPatch((x_pos, 0.72), 0.20, 0.24, boxstyle="round,pad=0.03", 
                                  fc='#0F172A', ec=col, lw=1.5, transform=ax2.transAxes)
    ax2.add_patch(rect)
    ax2.text(x_pos + 0.10, 0.87, title, color='#94A3B8', fontsize=8.5, ha='center', transform=ax2.transAxes)
    ax2.text(x_pos + 0.10, 0.76, val, color='#F8FAFC', fontsize=12, fontweight='bold', ha='center', transform=ax2.transAxes)

# Table headers & rows
col_labels = ["PID", "Arrival (AT)", "CPU1", "I/O Burst", "CPU2", "Completion (CT)", "TAT", "WT", "RT"]
table_vals = [
    ["P1", "0", "4", "2", "2", "19", "19", "11", "0"],
    ["P2", "1", "3", "1", "1", "21", "20", "15", "4"],
    ["P3", "2", "2", "0", "0", "11", "9", "7", "7"],
    ["P4", "3", "1", "0", "0", "13", "10", "9", "9"],
    ["P5", "4", "2", "1", "1", "23", "19", "15", "10"],
]

the_table = ax2.table(cellText=table_vals, colLabels=col_labels, loc='bottom', 
                      bbox=[0.05, 0.05, 0.90, 0.58], cellLoc='center')

the_table.auto_set_font_size(False)
the_table.set_fontsize(9)

for (r, c), cell in the_table.get_celld().items():
    if r == 0:
        cell.set_facecolor('#1E3A8A')
        cell.set_text_props(color='white', weight='bold')
    else:
        cell.set_facecolor('#1E293B' if r % 2 == 0 else '#0F172A')
        cell.set_text_props(color='#E2E8F0')
    cell.set_edgecolor('#334155')

plt.tight_layout()
sim_path = os.path.join(output_dir, "simulator_screenshot.png")
plt.savefig(sim_path, facecolor=fig.get_facecolor(), bbox_inches='tight')
plt.close()
print("Saved simulator screenshot:", sim_path)


# -------------------------------------------------------------
# 2. Slide 9: Professional Benchmark Bar Charts
# -------------------------------------------------------------
fig, (ax_lat, ax_eff) = plt.subplots(1, 2, figsize=(11, 4.6), dpi=200, gridspec_kw={'width_ratios': [1.3, 1]})
fig.patch.set_facecolor('#0B0F19')

# Policies & Metrics
policies = ['FCFS\n(Multi-Burst)', 'Priority\n(with Aging)', 'Round Robin\n(q=2, δ=1)', 'Round Robin\n(q=4, δ=1)']
avg_tat = [14.80, 12.40, 16.20, 14.60]
avg_wt  = [8.80,  6.40,  10.20, 8.60]
avg_rt  = [4.20,  1.80,  2.20,  3.40]
eff     = [85.71, 82.60, 73.08, 81.82]

x = np.arange(len(policies))
width = 0.25

# Left: Latency Metrics
ax_lat.set_facecolor('#111827')
bars1 = ax_lat.bar(x - width, avg_tat, width, label='Turnaround (TAT)', color='#38BDF8', edgecolor='#0B0F19')
bars2 = ax_lat.bar(x, avg_wt, width, label='Waiting (WT)', color='#F87171', edgecolor='#0B0F19')
bars3 = ax_lat.bar(x + width, avg_rt, width, label='Response (RT)', color='#FBBF24', edgecolor='#0B0F19')

ax_lat.set_title("Latency & Wait Time Comparison (Lower is Better)", fontsize=11.5, fontweight='bold', color='#F8FAFC', pad=10)
ax_lat.set_xticks(x)
ax_lat.set_xticklabels(policies, color='#CBD5E1', fontsize=8.5)
ax_lat.tick_params(colors='#94A3B8', labelsize=8.5)
ax_lat.set_ylabel("Time Units", color='#94A3B8', fontsize=9.5)
ax_lat.grid(axis='y', color='#1F2937', linestyle='--', alpha=0.7)
ax_lat.legend(facecolor='#1E293B', edgecolor='#374151', labelcolor='#F3F4F6', fontsize=8.5, loc='upper right')

# Value labels on bars
for bar in bars1:
    yval = bar.get_height()
    ax_lat.text(bar.get_x() + bar.get_width()/2, yval + 0.3, f"{yval:.1f}", ha='center', va='bottom', color='#38BDF8', fontsize=7.5, fontweight='bold')

for bar in bars2:
    yval = bar.get_height()
    ax_lat.text(bar.get_x() + bar.get_width()/2, yval + 0.3, f"{yval:.1f}", ha='center', va='bottom', color='#F87171', fontsize=7.5, fontweight='bold')

for spine in ax_lat.spines.values():
    spine.set_color('#374151')

# Right: CPU Efficiency
ax_eff.set_facecolor('#111827')
colors_eff = ['#60A5FA', '#34D399', '#F43F5E', '#10B981']
bars_eff = ax_eff.bar(policies, eff, width=0.45, color=colors_eff, edgecolor='#0B0F19')

ax_eff.set_title("CPU Efficiency η = ΣBT / Makespan (Higher is Better)", fontsize=11.5, fontweight='bold', color='#F8FAFC', pad=10)
ax_eff.set_xticklabels(policies, color='#CBD5E1', fontsize=8.5)
ax_eff.tick_params(colors='#94A3B8', labelsize=8.5)
ax_eff.set_ylabel("Efficiency Percentage (%)", color='#94A3B8', fontsize=9.5)
ax_eff.set_ylim(50, 100)
ax_eff.grid(axis='y', color='#1F2937', linestyle='--', alpha=0.7)

for bar in bars_eff:
    yval = bar.get_height()
    ax_eff.text(bar.get_x() + bar.get_width()/2, yval + 1.0, f"{yval:.1f}%", ha='center', va='bottom', color='white', fontsize=8.5, fontweight='bold')

for spine in ax_eff.spines.values():
    spine.set_color('#374151')

plt.tight_layout()
bench_path = os.path.join(output_dir, "benchmark_charts.png")
plt.savefig(bench_path, facecolor=fig.get_facecolor(), bbox_inches='tight')
plt.close()
print("Saved benchmark charts:", bench_path)
