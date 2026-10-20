"""Comprehensive benchmark runner for PyXL vs openpyxl in-memory operations (Phase 1).

Measures:
- Wall-clock time (median, mean, min, max, std dev)
- CPU user & sys time
- Peak RSS memory
- Throughput (cells/sec, rows/sec, ops/sec)
- Speedup and memory reduction ratio
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

REPETITIONS = 7
WARMUP_RUNS = 2

# Operations to benchmark
# 1. Workbook creation
# 2. Sequential cell assignment (10,000 cells: 1000 rows x 10 cols)
# 3. Random cell access (10,000 reads)
# 4. Row appending (5,000 rows x 10 cols = 50,000 cells)
# 5. Iteration over rows (5,000 rows x 10 cols, values_only=True)


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

        wb = Workbook()
        ws = wb.active
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c, value=r * 100 + c)

        c1 = time.process_time()
        t1 = time.perf_counter()
        rss = proc.memory_info().rss
        peak_rss = max(peak_rss, rss)
        wb.close()

        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    return {
        "operation": "Cell Assignment",
        "dataset": f"{rows}x{cols} ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_random_access(lib_name, rows=1000, cols=10, access_count=10000):
    import random
    mod = __import__(lib_name)
    Workbook = mod.Workbook

    wb = Workbook()
    ws = wb.active
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c, value=f"val_{r}_{c}")

    rng = random.Random(42)
    coords = [(rng.randint(1, rows), rng.randint(1, cols)) for _ in range(access_count)]

    # Warmup
    for r, c in coords[:100]:
        _ = ws.cell(row=r, column=c).value

    times = []
    cpu_times = []
    proc = psutil.Process()

    for _ in range(REPETITIONS):
        gc.collect()
        t0 = time.perf_counter()
        c0 = time.process_time()

        for r, c in coords:
            _ = ws.cell(row=r, column=c).value

        c1 = time.process_time()
        t1 = time.perf_counter()
        times.append(t1 - t0)
        cpu_times.append(c1 - c0)

    peak_rss = proc.memory_info().rss
    wb.close()

    return {
        "operation": "Random Cell Access",
        "dataset": f"{access_count:,} lookups",
        "total_units": access_count,
        "unit_name": "reads",
        "times": times,
        "cpu_times": cpu_times,
        "peak_rss_mb": round(peak_rss / (1024 * 1024), 2),
    }


def bench_row_append(lib_name, rows=5000, cols=10):
    mod = __import__(lib_name)
    Workbook = mod.Workbook
    total_cells = rows * cols
    sample_row = [f"item_{i}" for i in range(cols)]

    for _ in range(WARMUP_RUNS):
        wb = Workbook()
        ws = wb.active
        for _ in range(100):
            ws.append(sample_row)
        wb.close()

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

    # Warmup
    for _ in ws.iter_rows(values_only=True):
        pass

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
        "operation": "Row Iteration (values_only)",
        "dataset": f"{rows:,} rows ({total_cells:,} cells)",
        "total_units": total_cells,
        "unit_name": "cells",
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


def run_all():
    os.makedirs("benchmarks/reports", exist_ok=True)
    bench_fns = [
        bench_create_workbook,
        bench_cell_assignment,
        bench_random_access,
        bench_row_append,
        bench_iteration_values,
    ]

    results = []
    raw_data = {"environment": get_environment_info(), "benchmarks": []}

    print("================================================================================")
    print("PyXL vs openpyxl Benchmark Suite (Phase 1 — In-Memory Operations)")
    print("================================================================================")

    for fn in bench_fns:
        print(f"\nRunning {fn.__name__}...")
        pyxl_raw = fn("pyxl")
        openpyxl_raw = fn("openpyxl")

        pyxl_m = compute_metrics(pyxl_raw)
        openpyxl_m = compute_metrics(openpyxl_raw)

        speedup = openpyxl_m["median_sec"] / pyxl_m["median_sec"] if pyxl_m["median_sec"] > 0 else 1.0
        mem_diff_pct = (
            (1.0 - (pyxl_m["peak_rss_mb"] / openpyxl_m["peak_rss_mb"])) * 100
            if openpyxl_m["peak_rss_mb"] > 0
            else 0.0
        )

        row = {
            "Operation": pyxl_m["operation"],
            "Dataset": pyxl_m["dataset"],
            "openpyxl_time_s": round(openpyxl_m["median_sec"], 4),
            "pyxl_time_s": round(pyxl_m["median_sec"], 4),
            "speedup": round(speedup, 2),
            "openpyxl_cpu_s": round(openpyxl_m["median_cpu_sec"], 4),
            "pyxl_cpu_s": round(pyxl_m["median_cpu_sec"], 4),
            "openpyxl_rss_mb": openpyxl_m["peak_rss_mb"],
            "pyxl_rss_mb": pyxl_m["peak_rss_mb"],
            "mem_reduction_pct": round(mem_diff_pct, 1),
            "pyxl_throughput": f"{round(pyxl_m['throughput'], 1):,} {pyxl_m['unit_name']}/s",
            "openpyxl_throughput": f"{round(openpyxl_m['throughput'], 1):,} {openpyxl_m['unit_name']}/s",
        }
        results.append(row)

        raw_data["benchmarks"].append({
            "operation": pyxl_m["operation"],
            "dataset": pyxl_m["dataset"],
            "pyxl": pyxl_raw,
            "openpyxl": openpyxl_raw,
            "computed": row,
        })

        print(
            f"  -> {row['Operation']} [{row['Dataset']}]: "
            f"openpyxl={row['openpyxl_time_s']}s, pyxl={row['pyxl_time_s']}s, "
            f"Speedup={row['speedup']}x, MemReduction={row['mem_reduction_pct']}%"
        )

    # Save raw results JSON
    with open("benchmarks/reports/raw_results.json", "w") as f:
        json.dump(raw_data, f, indent=2)

    # Save environment JSON
    with open("benchmarks/reports/environment.json", "w") as f:
        json.dump(raw_data["environment"], f, indent=2)

    # Save summary CSV
    csv_file = "benchmarks/reports/summary.csv"
    with open(csv_file, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        writer.writeheader()
        writer.writerows(results)

    # Save comparison Markdown
    md_file = "benchmarks/reports/comparison.md"
    with open(md_file, "w") as f:
        f.write("# Phase 1 Comparative Benchmark Report: PyXL vs openpyxl\n\n")
        f.write(f"Generated at: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}\n\n")
        f.write(f"Environment: {raw_data['environment']['os']} ({raw_data['environment']['machine']}), ")
        f.write(f"{raw_data['environment']['cpu_count_logical']} Cores, {raw_data['environment']['total_ram_mb']} MB RAM\n\n")
        f.write("| Operation | Dataset | openpyxl (s) | PyXL (s) | Speedup | openpyxl RSS (MB) | PyXL RSS (MB) | Mem Reduction | PyXL Throughput |\n")
        f.write("|---|---|---|---|---|---|---|---|---|\n")
        for r in results:
            f.write(
                f"| {r['Operation']} | {r['Dataset']} | {r['openpyxl_time_s']}s | {r['pyxl_time_s']}s | "
                f"**{r['speedup']}x** | {r['openpyxl_rss_mb']} | {r['pyxl_rss_mb']} | {r['mem_reduction_pct']}% | {r['pyxl_throughput']} |\n"
            )

    print("\nBenchmark report successfully generated at benchmarks/reports/comparison.md")
    print("Summary CSV saved at benchmarks/reports/summary.csv")


if __name__ == "__main__":
    run_all()
