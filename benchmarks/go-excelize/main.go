package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"runtime/debug"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	defaultRows     = 200_000
	defaultColumns  = 20
	maxExcelRows    = 1_048_576
	maxExcelColumns = 16_384
	excelizeModule  = "github.com/xuri/excelize/v2"
)

type config struct {
	rows    int
	columns int
	output  string
}

type benchmarkResult struct {
	Implementation string  `json:"implementation"`
	LibraryVersion string  `json:"libraryVersion"`
	RuntimeVersion string  `json:"runtimeVersion"`
	Mode           string  `json:"mode"`
	WriteMode      string  `json:"writeMode"`
	Rows           int     `json:"rows"`
	Columns        int     `json:"columns"`
	Cells          int64   `json:"cells"`
	WriteMS        float64 `json:"writeMs"`
	SaveMS         float64 `json:"saveMs"`
	TotalMS        float64 `json:"totalMs"`
	MaxRSSBytes    *uint64 `json:"maxRssBytes"`
	OutputBytes    int64   `json:"outputBytes"`
}

func main() {
	cfg, err := parseConfig()
	if err == nil {
		err = run(cfg)
	}
	if err != nil {
		fmt.Fprintf(os.Stderr, "Go Excelize benchmark failed: %v\n", err)
		os.Exit(1)
	}
}

func parseConfig() (config, error) {
	rows := flag.Int("rows", defaultRows, "rows per workbook")
	columns := flag.Int("columns", defaultColumns, "columns per workbook")
	output := flag.String("output", "go-excelize-generation.xlsx", "output XLSX path")
	flag.Parse()

	if flag.NArg() != 0 {
		return config{}, fmt.Errorf("unexpected arguments: %v", flag.Args())
	}
	if *rows <= 0 || *rows > maxExcelRows {
		return config{}, fmt.Errorf("rows must be an integer from 1 to %d", maxExcelRows)
	}
	if *columns <= 0 || *columns > maxExcelColumns {
		return config{}, fmt.Errorf("columns must be an integer from 1 to %d", maxExcelColumns)
	}
	if *output == "" {
		return config{}, errors.New("output must not be empty")
	}

	return config{rows: *rows, columns: *columns, output: *output}, nil
}

func run(cfg config) error {
	output, err := filepath.Abs(cfg.output)
	if err != nil {
		return fmt.Errorf("resolve output path: %w", err)
	}
	if err := os.MkdirAll(filepath.Dir(output), 0o755); err != nil {
		return fmt.Errorf("create output directory: %w", err)
	}

	totalStart := time.Now()
	workbook := excelize.NewFile()
	closed := false
	defer func() {
		if !closed {
			_ = workbook.Close()
		}
	}()

	if err := workbook.SetSheetName("Sheet1", "Data"); err != nil {
		return fmt.Errorf("rename worksheet: %w", err)
	}
	stream, err := workbook.NewStreamWriter("Data")
	if err != nil {
		return fmt.Errorf("create stream writer: %w", err)
	}

	writeStart := time.Now()
	values := make([]interface{}, cfg.columns)
	for row := range cfg.rows {
		for column := range cfg.columns {
			if column%2 == 0 {
				values[column] = fmt.Sprintf("Row %d Col %d", row, column)
			} else {
				values[column] = float64(row) * float64(column) * 0.123
			}
		}
		if err := stream.SetRow("A"+strconv.Itoa(row+1), values); err != nil {
			return fmt.Errorf("write row %d: %w", row, err)
		}
	}
	if err := stream.Flush(); err != nil {
		return fmt.Errorf("flush stream: %w", err)
	}
	writeMS := millisecondsSince(writeStart)

	saveStart := time.Now()
	if err := workbook.SaveAs(output); err != nil {
		return fmt.Errorf("save workbook: %w", err)
	}
	if err := workbook.Close(); err != nil {
		return fmt.Errorf("close workbook: %w", err)
	}
	closed = true
	saveMS := millisecondsSince(saveStart)
	totalMS := millisecondsSince(totalStart)

	info, err := os.Stat(output)
	if err != nil {
		return fmt.Errorf("inspect output: %w", err)
	}

	result := benchmarkResult{
		Implementation: "go/excelize",
		LibraryVersion: dependencyVersion(excelizeModule),
		RuntimeVersion: runtime.Version(),
		Mode:           "streaming",
		WriteMode:      "row",
		Rows:           cfg.rows,
		Columns:        cfg.columns,
		Cells:          int64(cfg.rows) * int64(cfg.columns),
		WriteMS:        writeMS,
		SaveMS:         saveMS,
		TotalMS:        totalMS,
		MaxRSSBytes:    maxRSSBytes(),
		OutputBytes:    info.Size(),
	}
	if err := json.NewEncoder(os.Stdout).Encode(result); err != nil {
		return fmt.Errorf("encode result: %w", err)
	}
	return nil
}

func millisecondsSince(start time.Time) float64 {
	return float64(time.Since(start).Nanoseconds()) / 1_000_000
}

func dependencyVersion(module string) string {
	buildInfo, ok := debug.ReadBuildInfo()
	if !ok {
		return "unknown"
	}
	for _, dependency := range buildInfo.Deps {
		if dependency.Path == module {
			if dependency.Replace != nil {
				return dependency.Replace.Version
			}
			return dependency.Version
		}
	}
	return "unknown"
}
