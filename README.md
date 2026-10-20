# PyXL

A high-performance, Rust-powered alternative to [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl).

PyXL delivers an openpyxl-compatible Python API backed by a memory-safe, high-speed Rust engine built with PyO3 and Maturin.

## Key Features

- **openpyxl Drop-in Compatibility**: Designed to support standard openpyxl workflows seamlessly.
- **Rust Core**: Ultra-fast sparse cell storage and workbook memory management.
- **Bounded Memory & Streaming**: Minimal overhead and efficient memory layouts.
- **Strict Verification**: Thoroughly tested against openpyxl for 100% behavioral consistency.

## Phase Reports & Benchmarks

- [Phase 1 Implementation & Benchmark Report](docs/reports/phase1_report.md): In-Memory Core & PyO3 Bindings (up to 1,431× faster creation)
- [Phase 2 Implementation & Benchmark Report](docs/reports/phase2_report.md): High-Performance XLSX Reader (9.09× faster standard read, 552× faster time-to-first-row)
- [Phase 3 Implementation & Benchmark Report](docs/reports/phase3_report.md): High-Performance XLSX Writer (9.80× faster save, 7.85× faster streaming write-only, 236,988 rows/s)
- [Phase 4 Implementation Report](docs/reports/phase4_report.md): openpyxl Compatibility Layer & Priority API Parity
- [Phase 5 Performance & Benchmark Report](docs/reports/phase5_report.md): Multi-Scale Performance Optimizations (up to 8.12× faster, 1.68M cells/s)
- [API Compatibility Matrix](docs/compatibility_matrix.md): Exhaustive openpyxl vs PyXL feature matrix

## Quick Start

```python
import pyxl
from datetime import datetime

# Standard Workbook Creation & Saving (In-memory)
wb = pyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws["A1"] = "User"
ws["B1"] = "Amount"
ws.append(["Alice", 1500.50])
ws.append(["Bob", 2200.75])
wb.save("sales.xlsx")

# Streaming Write-Only Mode (True O(1) memory, high throughput)
wb = pyxl.Workbook(write_only=True)
ws = wb.create_sheet("StreamSheet")
for r in range(1, 100_001):
    ws.append([f"User_{r}", r * 10, r % 2 == 0, datetime.now()])
wb.save("stream_large.xlsx")

# Standard Reading (In-memory workbook)
wb = pyxl.load_workbook("sales.xlsx", data_only=True)
ws = wb.active
print(ws["A1"].value)

# Streaming Read-Only (Bounded O(1) memory, instant time-to-first-row)
with pyxl.load_workbook("stream_large.xlsx", read_only=True) as wb:
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        print(row)
```


## Development and Testing

Using Docker:

```bash
# Build test container
docker build -t pyxl-test .

# Run Rust tests
docker run --rm -v $(pwd):/workspace pyxl-test cargo test --workspace

# Run linter and formatting checks
docker run --rm -v $(pwd):/workspace pyxl-test cargo clippy --workspace --all-targets -- -D warnings
docker run --rm -v $(pwd):/workspace pyxl-test cargo fmt --all -- --check

# Run Python test suite
docker run --rm -v $(pwd):/workspace pyxl-test pytest tests/ -v

# Run comparative benchmarks
docker run --rm -v $(pwd):/workspace pyxl-test python benchmarks/run_benchmarks.py
```

## License

MIT OR Apache-2.0
