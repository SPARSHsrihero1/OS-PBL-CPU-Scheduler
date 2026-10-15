import copy
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

@dataclass
class Process:
    pid: str
    arrival_time: int
    cpu_burst_1: int
    io_burst: int = 0
    cpu_burst_2: int = 0
    priority: int = 0          # Lower number = Higher priority
    period: int = 0            # For Real-Time (RM/EDF)
    deadline: int = 0          # For Real-Time EDF
    
    # Execution states
    state: str = "NEW"         # NEW, READY, RUNNING, WAITING, TERMINATED
    current_phase: str = "CPU1"# CPU1, IO, CPU2, DONE
    rem_cpu1: int = 0
    rem_io: int = 0
    rem_cpu2: int = 0
    
    # Timing & Metrics
    start_time: int = -1
    completion_time: int = 0
    turnaround_time: int = 0
    waiting_time: int = 0
    response_time: int = -1
    total_cpu_time: int = 0
    
    def __post_init__(self):
        self.rem_cpu1 = self.cpu_burst_1
        self.rem_io = self.io_burst
        self.rem_cpu2 = self.cpu_burst_2
        self.total_cpu_time = self.cpu_burst_1 + self.cpu_burst_2

    def total_remaining_cpu(self) -> int:
        if self.current_phase == "CPU1":
            return self.rem_cpu1 + self.rem_cpu2
        elif self.current_phase == "IO":
            return self.rem_cpu2
        elif self.current_phase == "CPU2":
            return self.rem_cpu2
        return 0

@dataclass
class GanttEvent:
    label: str       # PID or "CS" (Context Switch) or "IDLE"
    start: int
    end: int
    event_type: str  # "PROCESS", "CONTEXT_SWITCH", "IDLE"

class OSKernelSimulator:

    @staticmethod
    def calculate_efficiency(total_cpu_burst: int, total_makespan: int) -> float:
        if total_makespan == 0:
            return 0.0
        return round((total_cpu_burst / total_makespan) * 100.0, 2)

    # STUDENT 1: FCFS with Multi-Burst & Context Switch Overhead (delta)
    @staticmethod
    def simulate_fcfs(processes: List[Process], cs_delta: int = 0):
        procs = [copy.deepcopy(p) for p in processes]
        ready_queue: List[Process] = []
        io_queue: List[Process] = []
        completed: List[Process] = []
        gantt: List[GanttEvent] = []

        time = 0
        current_running: Optional[Process] = None
        n = len(procs)

        while len(completed) < n:
            # 1. Admit new processes
            for p in procs:
                if p.arrival_time == time and p.state == "NEW":
                    p.state = "READY"
                    ready_queue.append(p)

            # 2. Progress I/O Queue
            newly_ready_from_io = []
            for p in io_queue:
                p.rem_io -= 1
                if p.rem_io == 0:
                    p.current_phase = "CPU2" if p.cpu_burst_2 > 0 else "DONE"
                    if p.current_phase == "DONE":
                        p.state = "TERMINATED"
                        p.completion_time = time
                        p.turnaround_time = p.completion_time - p.arrival_time
                        p.waiting_time = p.turnaround_time - p.total_cpu_time - p.io_burst
                        completed.append(p)
                    else:
                        p.state = "READY"
                        newly_ready_from_io.append(p)
            
            for p in newly_ready_from_io:
                io_queue.remove(p)
                ready_queue.append(p)

            # 3. CPU Dispatching
            if current_running is None:
                if ready_queue:
                    if cs_delta > 0 and (gantt and gantt[-1].label != "IDLE"):
                        gantt.append(GanttEvent("CS", time, time + cs_delta, "CONTEXT_SWITCH"))
                        time += cs_delta
                        for p in procs:
                            if p.arrival_time <= time and p.state == "NEW":
                                p.state = "READY"
                                ready_queue.append(p)

                    current_running = ready_queue.pop(0)
                    current_running.state = "RUNNING"
                    if current_running.response_time == -1:
                        current_running.response_time = time - current_running.arrival_time

            if current_running:
                if current_running.current_phase == "CPU1":
                    current_running.rem_cpu1 -= 1
                    phase_done = (current_running.rem_cpu1 == 0)
                else:
                    current_running.rem_cpu2 -= 1
                    phase_done = (current_running.rem_cpu2 == 0)

                if gantt and gantt[-1].label == current_running.pid and gantt[-1].end == time:
                    gantt[-1].end += 1
                else:
                    gantt.append(GanttEvent(current_running.pid, time, time + 1, "PROCESS"))

                time += 1

                if phase_done:
                    if current_running.current_phase == "CPU1" and current_running.io_burst > 0:
                        current_running.state = "WAITING"
                        current_running.current_phase = "IO"
                        io_queue.append(current_running)
                    else:
                        current_running.state = "TERMINATED"
                        current_running.completion_time = time
                        current_running.turnaround_time = current_running.completion_time - current_running.arrival_time
                        current_running.waiting_time = current_running.turnaround_time - current_running.total_cpu_time - current_running.io_burst
                        completed.append(current_running)
                    current_running = None
            else:
                if len(completed) < n:
                    if gantt and gantt[-1].label == "IDLE" and gantt[-1].end == time:
                        gantt[-1].end += 1
                    else:
                        gantt.append(GanttEvent("IDLE", time, time + 1, "IDLE"))
                    time += 1

        total_cpu = sum(p.total_cpu_time for p in procs)
        eff = OSKernelSimulator.calculate_efficiency(total_cpu, time)
        return gantt, completed, eff

    # STUDENT 2: Preemptive Priority with Dynamic Aging
    @staticmethod
    def simulate_priority(processes: List[Process], enable_aging: bool = True, aging_period: int = 5, cs_delta: int = 0):
        procs = [copy.deepcopy(p) for p in processes]
        gantt: List[GanttEvent] = []
        completed: List[Process] = []
        time = 0
        n = len(procs)
        dynamic_prio = {p.pid: p.priority for p in procs}
        last_pid = None

        while len(completed) < n:
            available = [p for p in procs if p.arrival_time <= time and p.total_remaining_cpu() > 0 and p.state != "TERMINATED"]
            if not available:
                if gantt and gantt[-1].label == "IDLE" and gantt[-1].end == time:
                    gantt[-1].end += 1
                else:
                    gantt.append(GanttEvent("IDLE", time, time + 1, "IDLE"))
                time += 1
                continue

            if enable_aging and time > 0 and time % aging_period == 0:
                for p in available:
                    dynamic_prio[p.pid] = max(0, dynamic_prio[p.pid] - 1)

            chosen = min(available, key=lambda x: (dynamic_prio[x.pid], x.arrival_time, x.pid))

            if last_pid is not None and last_pid != chosen.pid and cs_delta > 0:
                gantt.append(GanttEvent("CS", time, time + cs_delta, "CONTEXT_SWITCH"))
                time += cs_delta

            if chosen.response_time == -1:
                chosen.response_time = time - chosen.arrival_time

            if gantt and gantt[-1].label == chosen.pid and gantt[-1].end == time:
                gantt[-1].end += 1
            else:
                gantt.append(GanttEvent(chosen.pid, time, time + 1, "PROCESS"))

            last_pid = chosen.pid
            time += 1

            if chosen.current_phase == "CPU1":
                chosen.rem_cpu1 -= 1
                if chosen.rem_cpu1 == 0:
                    chosen.current_phase = "CPU2" if chosen.cpu_burst_2 > 0 else "DONE"
            elif chosen.current_phase == "CPU2":
                chosen.rem_cpu2 -= 1
                if chosen.rem_cpu2 == 0:
                    chosen.current_phase = "DONE"

            if chosen.total_remaining_cpu() == 0:
                chosen.state = "TERMINATED"
                chosen.completion_time = time
                chosen.turnaround_time = chosen.completion_time - chosen.arrival_time
                chosen.waiting_time = chosen.turnaround_time - chosen.total_cpu_time
                completed.append(chosen)

        total_cpu = sum(p.total_cpu_time for p in procs)
        eff = OSKernelSimulator.calculate_efficiency(total_cpu, time)
        return gantt, completed, eff

    # STUDENT 3: Round Robin with Context Switch Overhead (delta)
    @staticmethod
    def simulate_round_robin(processes: List[Process], quantum: int = 2, cs_delta: int = 0):
        procs = sorted([copy.deepcopy(p) for p in processes], key=lambda x: (x.arrival_time, x.pid))
        ready_queue: List[Process] = []
        completed: List[Process] = []
        gantt: List[GanttEvent] = []
        time = 0
        idx = 0
        n = len(procs)

        while idx < n and procs[idx].arrival_time <= time:
            ready_queue.append(procs[idx])
            idx += 1

        while ready_queue or idx < n:
            if not ready_queue:
                time = procs[idx].arrival_time
                while idx < n and procs[idx].arrival_time <= time:
                    ready_queue.append(procs[idx])
                    idx += 1
                continue

            curr = ready_queue.pop(0)

            if curr.response_time == -1:
                curr.response_time = time - curr.arrival_time

            exec_time = min(quantum, curr.total_remaining_cpu())

            if cs_delta > 0 and gantt and gantt[-1].label != "IDLE" and gantt[-1].label != curr.pid:
                gantt.append(GanttEvent("CS", time, time + cs_delta, "CONTEXT_SWITCH"))
                time += cs_delta

            gantt.append(GanttEvent(curr.pid, time, time + exec_time, "PROCESS"))
            time += exec_time

            rem_to_deduct = exec_time
            if curr.rem_cpu1 > 0:
                d = min(curr.rem_cpu1, rem_to_deduct)
                curr.rem_cpu1 -= d
                rem_to_deduct -= d
            if rem_to_deduct > 0 and curr.rem_cpu2 > 0:
                curr.rem_cpu2 -= rem_to_deduct

            while idx < n and procs[idx].arrival_time <= time:
                ready_queue.append(procs[idx])
                idx += 1

            if curr.total_remaining_cpu() > 0:
                ready_queue.append(curr)
            else:
                curr.state = "TERMINATED"
                curr.completion_time = time
                curr.turnaround_time = curr.completion_time - curr.arrival_time
                curr.waiting_time = curr.turnaround_time - curr.total_cpu_time
                completed.append(curr)

        total_cpu = sum(p.total_cpu_time for p in procs)
        eff = OSKernelSimulator.calculate_efficiency(total_cpu, time)
        return gantt, completed, eff

    # STUDENT 4: Real-Time Schedulers (Rate Monotonic & EDF)
    @staticmethod
    def real_time_feasibility_check(processes: List[Process]) -> Dict[str, any]:
        valid_procs = [p for p in processes if p.period > 0 and p.total_cpu_time > 0]
        n = len(valid_procs)
        if n == 0:
            return {"valid": False, "msg": "No periodic processes defined with Period > 0."}

        u = sum(p.total_cpu_time / p.period for p in valid_procs)
        rm_bound = n * ((2 ** (1.0 / n)) - 1)
        edf_bound = 1.00

        return {
            "valid": True,
            "utilization": round(u, 4),
            "rm_bound": round(rm_bound, 4),
            "edf_bound": edf_bound,
            "rm_schedulable": u <= rm_bound,
            "edf_schedulable": u <= edf_bound
        }
