# Phase 1 Comparative Benchmark Report: PyXL vs openpyxl

Generated at: 2026-10-20 10:16:42 UTC

Environment: Linux (aarch64), 8 Cores, 3916.36 MB RAM

| Operation | Dataset | openpyxl (s) | PyXL (s) | Speedup | openpyxl RSS (MB) | PyXL RSS (MB) | Mem Reduction | PyXL Throughput |
|---|---|---|---|---|---|---|---|---|
| Workbook Creation | 200 workbooks | 0.049s | 0.0s | **2009.78x** | 28.78 | 17.19 | 40.3% | 8,205,128.2 workbooks/s |
| Cell Assignment | 1000x10 (10,000 cells) | 0.0081s | 0.0031s | **2.59x** | 33.17 | 28.78 | 13.2% | 3,179,566.3 cells/s |
| Random Cell Access | 10,000 lookups | 0.0023s | 0.0036s | **0.62x** | 33.17 | 33.17 | 0.0% | 2,771,426.6 reads/s |
| Row Appending | 5,000 rows (50,000 cells) | 0.0566s | 0.0079s | **7.15x** | 52.68 | 36.83 | 30.1% | 632,364.4 rows/s |
| Row Iteration (values_only) | 5,000 rows (50,000 cells) | 0.0155s | 0.0037s | **4.23x** | 46.24 | 43.87 | 5.1% | 13,652,499.6 cells/s |
