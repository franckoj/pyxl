use rust_xlsxwriter::Workbook;
use std::env;
use std::fs;
use std::path::Path;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let rows: u32 = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(10_000);
    let cols: u16 = args.get(2).and_then(|s| s.parse().ok()).unwrap_or(10);
    let mode = args.get(3).map(|s| s.as_str()).unwrap_or("standard");

    let out_dir = Path::new("/tmp/rust_xlsxwriter_bench");
    fs::create_dir_all(out_dir)?;
    let out_path = out_dir.join(format!("bench_{}_{}_{}.xlsx", rows, cols, mode));

    let t0 = Instant::now();
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // standard mode
    // worksheet.dummy();

    for r in 0..rows {
        for c in 0..cols {
            worksheet.write_number(r, c, (r * 10 + c as u32) as f64)?;
        }
    }

    workbook.save(&out_path)?;
    let duration = t0.elapsed();

    let total_cells = (rows as u64) * (cols as u64);
    let throughput = (total_cells as f64) / duration.as_secs_f64();

    // Output JSON line for easy parsing
    println!(
        "{{\"rows\": {}, \"cols\": {}, \"cells\": {}, \"mode\": \"{}\", \"seconds\": {:.6}, \"throughput_cells_per_sec\": {:.0}}}",
        rows, cols, total_cells, mode, duration.as_secs_f64(), throughput
    );

    let _ = fs::remove_file(&out_path);
    Ok(())
}
