"""Comprehensive benchmark runner for PyXL vs openpyxl (Phase 1 & Phase 2).

Measures:
- Phase 1: In-Memory Operations (Workbook creation, cell access, append, iteration)
- Phase 2: XLSX File Reading (Standard loading, streaming read_only, time to first row, full iteration)
- Baselines: openpyxl (standard & read_only), pyxl (standard & read_only), python-calamine
- Metrics: Wall-clock time (median, mean, min, max, stdev), CPU time, Peak RSS, Throughput, Speedup
- Generates visual plots (PNG) and Markdown/CSV reports
"""

import csv
import gc
import json
import os
import platform
import statistics
import sys
import time
from datetime import datetime
import psutil
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

REPETITIONS = 7
WARMUP_RUNS = 2


def get_environment_info():
    uname = platform.uname()
    return {
        "timestamp": datetime.utcnow().isoformat(),
        "os": uname.system,
        "os_release": uname.release,
        "machine": uname.machine,
        "processor": uname.processor,
        "cpu_count_logical": psutil.cpu_count(logical=True),
        "cpu_count_physical": psutil.cpu_count(logical=False),
        "total_ram_mb": round(psutil.virtual_memory().total / (1024 * 1024), 2),
        "python_version": sys.version,
    }


def ensure_benchmark_file(rows=10000, cols=10):
    os.makedirs("benchmarks/data", exist_ok=True)
    filepath = f"benchmarks/data/bench_{rows // 1000}k.xlsx"
    if os.path.exists(filepath):
        return filepath

    print(f"Generating benchmark file with {rows:,} rows x {cols} cols: {filepath}...")
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DataSheet"

    headers = [f"Col_{c}" for c in range(1, cols + 1)]
    ws.append(headers)

    for r in range(1, rows + 1):
        row = [
            f"User_{r}",
            r * 10,
            r * 1.25,
            r % 2 == 0,
            datetime(2023, 1, 1, 12, 0, 0),
            f"Item_{r % 100}",
            r * 5,
            round(r * 0.05, 2),
            r % 3 == 0,
            "Active",
        ]
        ws.append(row)

    wb.save(filepath)
    print(f"Generated {filepath} ({round(os.path.getsize(filepath) / 1024, 1)} KB)")
    return filepath


# ==============================================================================
# Phase 1 Benchmarks (In-Memory)
# ==============================================================================

def bench_create_workbook(lib_name, iterations=200):
    mod = __import__(lib_name)
    Workbook = mod.Workbook

    for _ in range(WARMUP_RUNS):
        wb = Workbook()
        wb.close()

    times = []
    cpu_times = []
    mem_before = psutil.Process().memory_info().rss

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        for _ in range(iterations):
            wb = Workbook()
            wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mem_after = psutil.Process().memory_info().rss
    peak_rss = max(mem_after, mem_before)

    return {
        "operation": "Workbook Creation",
        "dataset": f"{iterations} workbooks",
        "total_units": iterations,
        "unit_name": "workbooks",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_cell_assignment(lib_name, rows=1000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols

    def run_once():
        wb = Workbook()
        ws = wb.active
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c, value=r * 100 + c)
        wb.close()

    for _ in range(WARMUP_RUNS):
        run_once()

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        run_once()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    return {
        "operation": "Cell Assignment",
        "dataset": f"{rows:,}x{cols} ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_row_append(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols
    sample_row = [100.5, "Alpha", True, 42, None, "Beta", 99.9, False, 12, "Gamma"]

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = Workbook()
        ws = wb.active
        for _ in range(rows):
            ws.append(sample_row)

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)
        wb.close()

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    return {
        "operation": "Row Appending",
        "dataset": f"{rows:,} rows ({total_cells:,} cells)",
        "total_units": rows,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_iteration_values(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols

    wb = Workbook()
    ws = wb.active
    sample_row = [100.5, "Alpha", True, 42, None, "Beta", 99.9, False, 12, "Gamma"]
    for _ in range(rows):
        ws.append(sample_row)

    times = []
    cpu_times = []
    proc = psutil.Process()

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        count = 0
        for row in ws.iter_rows(values_only=True):
            count += len(row)

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    peak_rss = proc.memory_info().rss
    wb.close()

    return {
        "operation": "In-Memory Row Iteration",
        "dataset": f"{rows:,} rows ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


# ==============================================================================
# Phase 2 Benchmarks (XLSX Reader)
# ==============================================================================

def bench_load_full_workbook(lib_name, filepath, read_only=False):
    mod = __import__(lib_name)
    load_workbook = mod.load_workbook

    # Warmup
    for _ in range(WARMUP_RUNS):
        wb = load_workbook(filepath, read_only=read_only)
        wb.close()

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = load_workbook(filepath, read_only=read_only)
        # Touch all rows
        count = 0
        for row in wb.active.iter_rows(values_only=True):
            count += len(row)
        wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mode = "read_only=True" if read_only else "Standard"
    return {
        "operation": f"Full Read ({mode})",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 10000,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_calamine_full_read(filepath):
    import python_calamine

    for _ in range(WARMUP_RUNS):
        wb = python_calamine.load_workbook(filepath)
        _ = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = python_calamine.load_workbook(filepath)
        sheet = wb.get_sheet_by_name(wb.sheet_names[0])
        rows = sheet.to_python()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    return {
        "operation": "Full Read (Calamine)",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 10000,
        "unit_name": "rows",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_time_to_first_row(lib_name, filepath, read_only=True):
    mod = __import__(lib_name)
    load_workbook = mod.load_workbook

    times = []
    cpu_times = []
    peak_rss = 0

    for _ in range(REPETITIONS):
        gc.collect()
        proc = psutil.Process()
        t0 = time.perf_counter()
        c0 = time.process_time()

        wb = load_workbook(filepath, read_only=read_only)
        row1 = next(iter(wb.active.iter_rows(values_only=True)))
        wb.close()

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    mode = "read_only=True" if read_only else "Standard"
    return {
        "operation": f"Time to First Row ({mode})",
        "dataset": "10,000 rows x 10 cols",
        "total_units": 1,
        "unit_name": "first_row",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def compute_metrics(bench_dict):
    times = bench_dict["times"]
    cpu_times = bench_dict["cpu_times"]
    total_units = bench_dict["total_units"]

    median_t = statistics.median(times)
    mean_t = statistics.mean(times)
    min_t = min(times)
    max_t = max(times)
    stdev_t = statistics.stdev(times) if len(times) > 1 else 0.0

    median_cpu = statistics.median(cpu_times)
    throughput = total_units / median_t if median_t > 0 else 0.0

    return {
        "operation": bench_dict["operation"],
        "dataset": bench_dict["dataset"],
        "unit_name": bench_dict["unit_name"],
        "median_sec": median_t,
        "mean_sec": mean_t,
        "min_sec": min_t,
        "max_sec": max_t,
        "stdev_sec": stdev_t,
        "median_cpu_sec": median_cpu,
        "peak_rss_mb": bench_dict["peak_rss_mb"],
        "throughput": throughput,
    }


def generate_phase2_plots(p2_data):
    os.makedirs("benchmarks/reports", exist_ok=True)

    # 1. Reading Execution Time Comparison
    ops = [item["name"] for item in p2_data]
    times = [item["time_s"] for item in p2_data]
    colors = ["#e74c3c", "#f39c12", "#2ecc71", "#1abc9c", "#3498db"]

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(ops, times, color=colors[:len(ops)], width=0.5)
    ax.set_ylabel("Execution Time (seconds) - Lower is better")
    ax.set_title("XLSX Read Time (10,000 rows x 10 columns)")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.4f}s",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_reading_time.png", dpi=150)
    plt.close()

    # 2. Reading Speedup Comparison vs openpyxl baseline
    openpyxl_std_time = next(item["time_s"] for item in p2_data if "openpyxl (Standard)" in item["name"])
    speedup_items = []
    for item in p2_data:
        speedup = openpyxl_std_time / item["time_s"] if item["time_s"] > 0 else 1.0
        speedup_items.append((item["name"], speedup))

    fig, ax = plt.subplots(figsize=(10, 5))
    s_names = [x[0] for x in speedup_items]
    s_vals = [x[1] for x in speedup_items]
    bars = ax.bar(s_names, s_vals, color="#2980b9", width=0.5)
    ax.axhline(1.0, color="gray", linestyle="--", linewidth=1, label="Baseline (openpyxl std = 1.0x)")
    ax.set_ylabel("Speedup Multiplier vs openpyxl (Higher is better)")
    ax.set_title("XLSX Reading Speedup Multiplier vs openpyxl Standard")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f}x",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_speedup.png", dpi=150)
    plt.close()

    # 3. Peak RSS Memory
    mems = [item["rss_mb"] for item in p2_data]
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(ops, mems, color=["#e67e22", "#d35400", "#27ae60", "#16a085", "#2980b9"][:len(ops)], width=0.5)
    ax.set_ylabel("Peak RSS Memory (MB) - Lower is better")
    ax.set_title("Peak Memory Usage During XLSX Read (10,000 rows)")
    for bar in bars:
        h = bar.get_height()
        ax.annotate(f"{h:.1f} MB",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    plt.savefig("benchmarks/reports/phase2_memory.png", dpi=150)
    plt.close()


def run_all():
    os.makedirs("benchmarks/reports", exist_ok=True)
    os.makedirs("docs/reports", exist_ok=True)
    bench_file = ensure_benchmark_file(rows=10000, cols=10)

    print("================================================================================")
    print("PyXL Comprehensive Benchmark Suite (Phase 1 In-Memory & Phase 2 XLSX Reader)")
    print("================================================================================")

    # -------------------------------------------------------------------------
    # Part 1: Phase 1 In-Memory Operations
    # -------------------------------------------------------------------------
    print("\n>>> Running Phase 1 In-Memory Operations...")
    p1_fns = [
        bench_create_workbook,
        bench_cell_assignment,
        bench_row_append,
        bench_iteration_values,
    ]
    p1_results = []
    for fn in p1_fns:
        print(f"  Testing {fn.__name__}...")
        pyxl_m = compute_metrics(fn("pyxl"))
        openpyxl_m = compute_metrics(fn("openpyxl"))

        speedup = openpyxl_m["median_sec"] / pyxl_m["median_sec"] if pyxl_m["median_sec"] > 0 else 1.0
        mem_diff_pct = (
            (1.0 - (pyxl_m["peak_rss_mb"] / openpyxl_m["peak_rss_mb"])) * 100
            if openpyxl_m["peak_rss_mb"] > 0
            else 0.0
        )

        p1_results.append({
            "Operation": pyxl_m["operation"],
            "Dataset": pyxl_m["dataset"],
            "openpyxl_time_s": round(openpyxl_m["median_sec"], 4),
            "pyxl_time_s": round(pyxl_m["median_sec"], 4),
            "speedup": round(speedup, 2),
            "openpyxl_rss_mb": openpyxl_m["peak_rss_mb"],
            "pyxl_rss_mb": pyxl_m["peak_rss_mb"],
            "mem_reduction_pct": round(mem_diff_pct, 1),
            "pyxl_throughput": f"{round(pyxl_m['throughput'], 1):,} {pyxl_m['unit_name']}/s",
            "openpyxl_throughput": f"{round(openpyxl_m['throughput'], 1):,} {openpyxl_m['unit_name']}/s",
        })

    # -------------------------------------------------------------------------
    # Part 2: Phase 2 XLSX Reader Operations
    # -------------------------------------------------------------------------
    print("\n>>> Running Phase 2 XLSX Reader Benchmarks (10,000 rows x 10 cols)...")
    m_op_std = compute_metrics(bench_load_full_workbook("openpyxl", bench_file, read_only=False))
    m_op_ro = compute_metrics(bench_load_full_workbook("openpyxl", bench_file, read_only=True))
    m_px_std = compute_metrics(bench_load_full_workbook("pyxl", bench_file, read_only=False))
    m_px_ro = compute_metrics(bench_load_full_workbook("pyxl", bench_file, read_only=True))
    m_cal = compute_metrics(bench_calamine_full_read(bench_file))

    p2_chart_data = [
        {"name": "openpyxl (Standard)", "time_s": round(m_op_std["median_sec"], 4), "rss_mb": m_op_std["peak_rss_mb"]},
        {"name": "openpyxl (read_only)", "time_s": round(m_op_ro["median_sec"], 4), "rss_mb": m_op_ro["peak_rss_mb"]},
        {"name": "PyXL (Standard)", "time_s": round(m_px_std["median_sec"], 4), "rss_mb": m_px_std["peak_rss_mb"]},
        {"name": "PyXL (read_only)", "time_s": round(m_px_ro["median_sec"], 4), "rss_mb": m_px_ro["peak_rss_mb"]},
        {"name": "calamine (C-Rust)", "time_s": round(m_cal["median_sec"], 4), "rss_mb": m_cal["peak_rss_mb"]},
    ]

    # Time to first row
    ttfr_op_std = compute_metrics(bench_time_to_first_row("openpyxl", bench_file, read_only=False))
    ttfr_op_ro = compute_metrics(bench_time_to_first_row("openpyxl", bench_file, read_only=True))
    ttfr_px_std = compute_metrics(bench_time_to_first_row("pyxl", bench_file, read_only=False))
    ttfr_px_ro = compute_metrics(bench_time_to_first_row("pyxl", bench_file, read_only=True))

    p2_results = [
        {
            "Operation": "Full XLSX Read (Standard)",
            "openpyxl": f"{m_op_std['median_sec']:.4f}s",
            "pyxl": f"{m_px_std['median_sec']:.4f}s",
            "speedup": f"{m_op_std['median_sec'] / m_px_std['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_std['peak_rss_mb']:.1f} MB",
            "pyxl_rss": f"{m_px_std['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_px_std['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Full XLSX Read (read_only=True)",
            "openpyxl": f"{m_op_ro['median_sec']:.4f}s",
            "pyxl": f"{m_px_ro['median_sec']:.4f}s",
            "speedup": f"{m_op_ro['median_sec'] / m_px_ro['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_ro['peak_rss_mb']:.1f} MB",
            "pyxl_rss": f"{m_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_px_ro['throughput']:,.0f} rows/s",
        },
        {
            "Operation": "Time to First Row (read_only=True)",
            "openpyxl": f"{ttfr_op_ro['median_sec']*1000:.2f} ms",
            "pyxl": f"{ttfr_px_ro['median_sec']*1000:.2f} ms",
            "speedup": f"{ttfr_op_ro['median_sec'] / ttfr_px_ro['median_sec']:.2f}x",
            "openpyxl_rss": f"{ttfr_op_ro['peak_rss_mb']:.1f} MB",
            "pyxl_rss": f"{ttfr_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": "Instantaneous",
        },
        {
            "Operation": "Time to First Row (Standard vs RO)",
            "openpyxl": f"{ttfr_op_std['median_sec']*1000:.2f} ms",
            "pyxl": f"{ttfr_px_ro['median_sec']*1000:.2f} ms",
            "speedup": f"{ttfr_op_std['median_sec'] / ttfr_px_ro['median_sec']:.1f}x",
            "openpyxl_rss": f"{ttfr_op_std['peak_rss_mb']:.1f} MB",
            "pyxl_rss": f"{ttfr_px_ro['peak_rss_mb']:.1f} MB",
            "throughput": "Instantaneous",
        },
        {
            "Operation": "Calamine Full Read Baseline",
            "openpyxl": f"{m_op_std['median_sec']:.4f}s (openpyxl std)",
            "pyxl": f"{m_cal['median_sec']:.4f}s (calamine)",
            "speedup": f"{m_op_std['median_sec'] / m_cal['median_sec']:.2f}x",
            "openpyxl_rss": f"{m_op_std['peak_rss_mb']:.1f} MB",
            "pyxl_rss": f"{m_cal['peak_rss_mb']:.1f} MB",
            "throughput": f"{m_cal['throughput']:,.0f} rows/s",
        },
    ]

    # Generate plots
    generate_phase2_plots(p2_chart_data)

    # Save summary Markdown
    md_file = "benchmarks/reports/comparison.md"
    env = get_environment_info()
    with open(md_file, "w") as f:
        f.write("# PyXL Comprehensive Benchmark Report (Phase 1 & Phase 2)\n\n")
        f.write(f"Generated at: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}\n\n")
        f.write(f"**Environment**: {env['os']} ({env['machine']}), {env['cpu_count_logical']} Cores, {env['total_ram_mb']} MB RAM\n\n")

        f.write("## 1. Phase 2: High-Performance XLSX Reader Benchmarks\n\n")
        f.write("Dataset: `10,000 rows x 10 columns` (100,000 cells) containing strings, floats, ints, bools, and datetimes.\n\n")
        f.write("| Operation | openpyxl | PyXL | Speedup | openpyxl RSS | PyXL RSS | PyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|\n")
        for r in p2_results:
            f.write(f"| {r['Operation']} | {r['openpyxl']} | {r['pyxl']} | **{r['speedup']}** | {r['openpyxl_rss']} | {r['pyxl_rss']} | {r['throughput']} |\n")

        f.write("\n### Phase 2 Performance Graphs\n\n")
        f.write("#### 1. Total Reading Execution Time (10,000 rows)\n")
        f.write("![Reading Time](phase2_reading_time.png)\n\n")
        f.write("#### 2. Speedup Factor over openpyxl\n")
        f.write("![Reading Speedup](phase2_speedup.png)\n\n")
        f.write("#### 3. Peak RSS Memory Utilization\n")
        f.write("![Reading Memory](phase2_memory.png)\n\n")

        f.write("## 2. Phase 1: In-Memory Operations Benchmarks\n\n")
        f.write("| Operation | Dataset | openpyxl (s) | PyXL (s) | Speedup | openpyxl RSS | PyXL RSS | Mem Reduction | PyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|---|---|\n")
        for r in p1_results:
            f.write(f"| {r['Operation']} | {r['Dataset']} | {r['openpyxl_time_s']}s | {r['pyxl_time_s']}s | **{r['speedup']}x** | {r['openpyxl_rss_mb']} MB | {r['pyxl_rss_mb']} MB | {r['mem_reduction_pct']}% | {r['pyxl_throughput']} |\n")

    # Save summary JSON
    results_json = {
        "environment": env,
        "phase1": p1_results,
        "phase2": p2_results,
        "phase2_chart_data": p2_chart_data,
    }
    with open("benchmarks/reports/phase2_results.json", "w") as f:
        json.dump(results_json, f, indent=2)

    print("\n================================================================================")
    print("Benchmark completed successfully!")
    print(f"Report: {md_file}")
    print("JSON: benchmarks/reports/phase2_results.json")
    print("Plots: benchmarks/reports/phase2_*.png")
    print("================================================================================")


if __name__ == "__main__":
    run_all()
