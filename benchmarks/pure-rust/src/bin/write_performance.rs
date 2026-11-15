//! The write benchmark's Rust side, without Criterion, so that
//! `benchmarks/node-write-performance.js` can time the same workloads through
//! the Node.js package, ExcelJS and SheetJS with the same sampling.
//!
//! The `write-performance` workload is rust_xlsxwriter's own benchmark: its
//! configurations, sample sizes and `write_*` functions are copied from
//! <https://github.com/jmcnamara/rust_xlsxwriter/blob/v0.102.0/benches/write_performance.rs>,
//! Copyright 2022-2026, John McNamara, jmcnamara@cpan.org (MIT OR Apache-2.0).
//! The `export` workload is the business-report export of
//! [`rusc_xlsx_pure_rust_benchmark::export`], `--rows` rows of 30 columns.
//!
//! Upstream selects the memory mode at compile time; this binary takes it as
//! `--mode` because the workspace always compiles the `constant_memory`
//! feature, as the package does. Every configuration prints one JSON line
//! with the time of each measured iteration. With `--validate-dir DIR` it
//! saves each workbook once to DIR instead of timing it, so the runner can
//! check that the implementations write the same data, and `--memory GROUP`
//! writes the largest workbook of GROUP once and prints the peak RSS.

use std::env;
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use rusc_xlsx_pure_rust_benchmark::{export, max_rss_bytes};
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

// All configurations use a fixed column count and vary the row count.
const COLS: u16 = 50;

// Rows for 1K, 10K, 100K and 1M cells.
const CONFIGS: [u32; 4] = [20, 200, 2_000, 20_000];

// The configuration whose memory use is measured.
const MEMORY_ROWS: u32 = 20_000;

type Populate = fn(&mut Worksheet, u32, u16) -> Result<(), XlsxError>;

// Upstream's `criterion_group!` order.
const GROUPS: [(&str, Populate); 4] = [
    ("numbers", write_numbers),
    ("mixed", write_mixed),
    ("string_static", write_string_static),
    ("string_dynamic", write_string_dynamic),
];

const DEFAULT_WARM_UP_MS: u64 = 1_000;

/// The columns, row counts, measured row count and groups of a workload.
struct Workload {
    columns: u16,
    configs: Vec<u32>,
    memory_rows: u32,
    groups: Vec<(&'static str, Populate)>,
}

impl Workload {
    fn parse(name: &str, rows: Option<u32>) -> Result<Self, String> {
        match (name, rows) {
            ("write-performance", None) => Ok(Self {
                columns: COLS,
                configs: CONFIGS.to_vec(),
                memory_rows: MEMORY_ROWS,
                groups: GROUPS.to_vec(),
            }),
            ("write-performance", Some(_)) => {
                Err("--rows applies to the export workload only".to_owned())
            }
            ("export", Some(rows)) => Ok(Self {
                columns: export::COLUMNS,
                configs: vec![rows],
                memory_rows: rows,
                groups: vec![("export", export::write_export)],
            }),
            ("export", None) => Err("the export workload needs --rows".to_owned()),
            _ => Err(format!(
                "workload must be one of: write-performance, export (got {name})"
            )),
        }
    }
}

#[derive(Clone, Copy)]
enum Mode {
    Default,
    Constant,
    Low,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "default" => Ok(Self::Default),
            "constant" => Ok(Self::Constant),
            "low" => Ok(Self::Low),
            _ => Err(format!(
                "mode must be one of: default, constant, low (got {value})"
            )),
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Constant => "constant",
            Self::Low => "low",
        }
    }
}

struct Config {
    workload: Workload,
    mode: Mode,
    warm_up: Duration,
    validate_dir: Option<PathBuf>,
    memory_group: Option<(&'static str, Populate)>,
}

fn main() {
    let result = parse_args().and_then(|config| {
        run(&config).map_err(|error| format!("Pure Rust write_performance failed: {error}"))
    });

    if let Err(message) = result {
        eprintln!("{message}");
        std::process::exit(1);
    }
}

fn parse_args() -> Result<Config, String> {
    let args: Vec<String> = env::args().skip(1).collect();
    let mut workload_name = "write-performance".to_owned();
    let mut rows = None;
    let mut mode = Mode::Default;
    let mut warm_up = Duration::from_millis(DEFAULT_WARM_UP_MS);
    let mut validate_dir = None;
    let mut memory_group_name = None;

    for pair in args.chunks(2) {
        let [flag, value] = pair else {
            return Err(format!("Missing value for {}", pair[0]));
        };
        match flag.as_str() {
            "--workload" => workload_name.clone_from(value),
            "--rows" => {
                let parsed = value
                    .parse::<u32>()
                    .ok()
                    .filter(|&rows| rows > 0)
                    .ok_or_else(|| "rows must be a positive integer".to_owned())?;
                rows = Some(parsed);
            }
            "--mode" => mode = Mode::parse(value)?,
            "--warm-up-ms" => {
                let millis = value
                    .parse::<u64>()
                    .map_err(|_| "warm-up-ms must be a non-negative integer".to_owned())?;
                warm_up = Duration::from_millis(millis);
            }
            "--validate-dir" => validate_dir = Some(PathBuf::from(value)),
            "--memory" => memory_group_name = Some(value.clone()),
            _ => return Err(format!("Unknown argument: {flag}")),
        }
    }

    let workload = Workload::parse(&workload_name, rows)?;
    let memory_group = memory_group_name
        .map(|name| {
            workload
                .groups
                .iter()
                .copied()
                .find(|(group, _)| *group == name)
                .ok_or_else(|| format!("Unknown group: {name}"))
        })
        .transpose()?;

    Ok(Config {
        workload,
        mode,
        warm_up,
        validate_dir,
        memory_group,
    })
}

fn run(config: &Config) -> Result<(), XlsxError> {
    let workload = &config.workload;
    let cols = workload.columns;

    if let Some((group, populate)) = config.memory_group {
        return measure_memory(config.mode, workload, group, populate);
    }

    for &(group, populate) in &workload.groups {
        for &rows in &workload.configs {
            let cell_count = rows * u32::from(cols);

            if let Some(dir) = &config.validate_dir {
                let path = dir.join(format!("{group}-{cell_count}.xlsx"));
                save_workbook(rows, cols, config.mode, populate, &path)?;
                continue;
            }

            // At least one warm-up iteration, then as many as fit the warm-up time.
            let warm_up_start = Instant::now();
            loop {
                write_workbook(black_box(rows), black_box(cols), config.mode, populate)?;
                if warm_up_start.elapsed() >= config.warm_up {
                    break;
                }
            }

            let mut samples = Vec::with_capacity(sample_size(cell_count));
            for _ in 0..sample_size(cell_count) {
                let start = Instant::now();
                write_workbook(black_box(rows), black_box(cols), config.mode, populate)?;
                samples.push(start.elapsed().as_secs_f64() * 1_000.0);
            }

            let samples = samples
                .iter()
                .map(|millis| format!("{millis:.6}"))
                .collect::<Vec<_>>()
                .join(",");
            println!(
                concat!(
                    "{{\"implementation\":\"pure-rust\",\"mode\":\"{}\",",
                    "\"group\":\"{}\",\"rows\":{},\"columns\":{},",
                    "\"cells\":{},\"samplesMs\":[{}]}}"
                ),
                config.mode.name(),
                group,
                rows,
                cols,
                cell_count,
                samples
            );
        }
    }

    Ok(())
}

/// Write the largest workbook of a group once and print the peak RSS before
/// and after it.
fn measure_memory(
    mode: Mode,
    workload: &Workload,
    group: &str,
    populate: Populate,
) -> Result<(), XlsxError> {
    let rows = workload.memory_rows;
    let cols = workload.columns;
    let baseline = max_rss_bytes();
    let temp_file =
        std::env::temp_dir().join(format!("bench_write_perf_{}.xlsx", std::process::id()));
    save_workbook(rows, cols, mode, populate, &temp_file)?;
    let peak = max_rss_bytes();
    let output_bytes = std::fs::metadata(&temp_file)?.len();
    let _ = std::fs::remove_file(temp_file);

    let json = |bytes: Option<u64>| bytes.map_or_else(|| "null".to_owned(), |b| b.to_string());
    println!(
        concat!(
            "{{\"implementation\":\"pure-rust\",\"mode\":\"{}\",\"group\":\"{}\",\"rows\":{},",
            "\"cells\":{},\"baselineRssBytes\":{},\"peakRssBytes\":{},\"outputBytes\":{}}}"
        ),
        mode.name(),
        group,
        rows,
        rows * u32::from(cols),
        json(baseline),
        json(peak),
        output_bytes
    );

    Ok(())
}

/// Fewer samples for the larger, slower configs.
fn sample_size(cell_count: u32) -> usize {
    if cell_count >= 1_000_000 {
        10
    } else if cell_count >= 100_000 {
        20
    } else {
        50
    }
}

/// Create a workbook, populate it and save it to a temporary file.
fn write_workbook(rows: u32, cols: u16, mode: Mode, populate: Populate) -> Result<(), XlsxError> {
    let temp_file =
        std::env::temp_dir().join(format!("bench_write_perf_{}.xlsx", std::process::id()));
    save_workbook(rows, cols, mode, populate, &temp_file)?;
    let _ = std::fs::remove_file(temp_file);

    Ok(())
}

fn save_workbook(
    rows: u32,
    cols: u16,
    mode: Mode,
    populate: Populate,
    path: &Path,
) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    populate(add_benchmark_worksheet(&mut workbook, mode)?, rows, cols)?;
    workbook.save(path)
}

/// Add a worksheet in the memory mode selected by `--mode`.
fn add_benchmark_worksheet(
    workbook: &mut Workbook,
    mode: Mode,
) -> Result<&mut Worksheet, XlsxError> {
    match mode {
        Mode::Default => Ok(workbook.add_worksheet()),
        Mode::Constant => workbook.add_worksheet_with_constant_memory(),
        Mode::Low => workbook.add_worksheet_with_low_memory(),
    }
}

/// Write numbers and strings, alternatively, to each cell in the range.
fn write_mixed(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            if col % 2 == 1 {
                worksheet.write_string(row, col, "Foo")?;
            } else {
                worksheet.write_number(row, col, 12345.0)?;
            }
        }
    }

    Ok(())
}

/// Write the same number to every cell.
fn write_numbers(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_number(row, col, 12345.0)?;
        }
    }

    Ok(())
}

/// Write the same static string to every cell.
fn write_string_static(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_string(row, col, "Foo")?;
        }
    }

    Ok(())
}

/// Write a dynamic string to every cell.
fn write_string_dynamic(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    for row in 0..rows {
        for col in 0..cols {
            worksheet.write_string(row, col, format!("Cell ({row}, {col})"))?;
        }
    }

    Ok(())
}
