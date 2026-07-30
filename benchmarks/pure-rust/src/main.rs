use rust_xlsxwriter::Workbook;
use std::env;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;

const DEFAULT_ROWS: u32 = 200_000;
const DEFAULT_COLUMNS: u16 = 20;
const MAX_EXCEL_ROWS: u32 = 1_048_576;
const MAX_EXCEL_COLUMNS: u16 = 16_384;

struct Config {
    rows: u32,
    columns: u16,
    output: PathBuf,
}

fn main() {
    match parse_args().and_then(run) {
        Ok(()) => {}
        Err(message) => {
            eprintln!("{message}");
            std::process::exit(1);
        }
    }
}

fn parse_args() -> Result<Config, String> {
    let args: Vec<String> = env::args().skip(1).collect();

    if args.iter().any(|arg| arg == "--help" || arg == "-h") {
        println!("Usage: rusc-xlsx-pure-rust-benchmark [--rows N] [--columns N] [--output PATH]");
        std::process::exit(0);
    }

    let mut rows = DEFAULT_ROWS;
    let mut columns = DEFAULT_COLUMNS;
    let mut output = PathBuf::from("pure-rust-generation.xlsx");
    let mut index = 0;

    while index < args.len() {
        let flag = &args[index];
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("Missing value for {flag}"))?;

        match flag.as_str() {
            "--rows" => {
                rows = parse_positive_u32(value, "rows")?;
                if rows > MAX_EXCEL_ROWS {
                    return Err(format!(
                        "rows must not exceed Excel's limit of {MAX_EXCEL_ROWS}"
                    ));
                }
            }
            "--columns" => {
                let parsed = parse_positive_u32(value, "columns")?;
                if parsed > u32::from(MAX_EXCEL_COLUMNS) {
                    return Err(format!(
                        "columns must not exceed Excel's limit of {MAX_EXCEL_COLUMNS}"
                    ));
                }
                columns = parsed as u16;
            }
            "--output" => output = PathBuf::from(value),
            _ => return Err(format!("Unknown argument: {flag}")),
        }

        index += 2;
    }

    Ok(Config {
        rows,
        columns,
        output,
    })
}

fn parse_positive_u32(value: &str, name: &str) -> Result<u32, String> {
    let parsed = value
        .parse::<u32>()
        .map_err(|_| format!("{name} must be a positive integer"))?;

    if parsed == 0 {
        return Err(format!("{name} must be a positive integer"));
    }

    Ok(parsed)
}

fn run(config: Config) -> Result<(), String> {
    generate(&config).map_err(|error| format!("Pure Rust benchmark failed: {error}"))
}

fn generate(config: &Config) -> Result<(), Box<dyn Error>> {
    let output_dir = config
        .output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(output_dir)?;

    let total_start = Instant::now();
    let mut workbook = Workbook::new();
    workbook.set_tempdir(output_dir)?;
    let worksheet = workbook.add_worksheet_with_constant_memory();
    worksheet.set_name("Data")?;

    let write_start = Instant::now();
    for row in 0..config.rows {
        for column in 0..config.columns {
            if column % 2 == 0 {
                worksheet.write_string(row, column, format!("Row {row} Col {column}"))?;
            } else {
                let value = f64::from(row) * f64::from(column) * 0.123;
                worksheet.write_number(row, column, value)?;
            }
        }
    }
    let write_ms = write_start.elapsed().as_secs_f64() * 1_000.0;

    let save_start = Instant::now();
    workbook.save(&config.output)?;
    let save_ms = save_start.elapsed().as_secs_f64() * 1_000.0;
    let total_ms = total_start.elapsed().as_secs_f64() * 1_000.0;

    let output_bytes = fs::metadata(&config.output)?.len();
    let cells = u64::from(config.rows) * u64::from(config.columns);
    let max_rss = max_rss_bytes()
        .map(|bytes| bytes.to_string())
        .unwrap_or_else(|| "null".to_owned());

    println!(
        concat!(
            "{{\"implementation\":\"pure-rust\",",
            "\"mode\":\"constant-memory\",",
            "\"rows\":{},\"columns\":{},\"cells\":{},",
            "\"writeMs\":{:.3},\"saveMs\":{:.3},\"totalMs\":{:.3},",
            "\"maxRssBytes\":{},\"outputBytes\":{}}}"
        ),
        config.rows, config.columns, cells, write_ms, save_ms, total_ms, max_rss, output_bytes
    );

    Ok(())
}

#[cfg(unix)]
fn max_rss_bytes() -> Option<u64> {
    let mut usage = std::mem::MaybeUninit::<libc::rusage>::zeroed();
    let status = unsafe { libc::getrusage(libc::RUSAGE_SELF, usage.as_mut_ptr()) };
    if status != 0 {
        return None;
    }

    let max_rss = unsafe { usage.assume_init() }.ru_maxrss;
    if max_rss < 0 {
        return None;
    }

    let max_rss = max_rss as u64;
    if cfg!(any(target_os = "macos", target_os = "ios")) {
        Some(max_rss)
    } else {
        Some(max_rss.saturating_mul(1_024))
    }
}

#[cfg(not(unix))]
fn max_rss_bytes() -> Option<u64> {
    None
}
