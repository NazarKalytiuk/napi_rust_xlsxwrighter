using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal static class Program
{
    private const int DefaultRows = 200_000;
    private const int DefaultColumns = 20;
    private const int MaxExcelRows = 1_048_576;
    private const int MaxExcelColumns = 16_384;
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private sealed record Config(int Rows, int Columns, string Output);

    [StructLayout(LayoutKind.Explicit, Size = 256)]
    private struct ResourceUsage
    {
        [FieldOffset(32)]
        public long MaxRss;
    }

    [DllImport("libc", EntryPoint = "getrusage", SetLastError = true)]
    private static extern int GetResourceUsage(int target, out ResourceUsage usage);

    public static int Main(string[] args)
    {
        try
        {
            if (args.Any(argument => argument is "--help" or "-h"))
            {
                Console.WriteLine(
                    "Usage: rusc-xlsx-csharp-benchmark [--rows N] [--columns N] [--output PATH]"
                );
                return 0;
            }

            Run(ParseArgs(args));
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"C# Open XML SDK benchmark failed: {error.Message}");
            return 1;
        }
    }

    private static Config ParseArgs(string[] args)
    {
        var rows = DefaultRows;
        var columns = DefaultColumns;
        var output = "csharp-openxml-generation.xlsx";

        for (var index = 0; index < args.Length; index += 2)
        {
            var flag = args[index];
            if (index + 1 >= args.Length)
            {
                throw new ArgumentException($"Missing value for {flag}");
            }

            var value = args[index + 1];
            switch (flag)
            {
                case "--rows":
                    rows = ParsePositiveInteger(value, "rows", MaxExcelRows);
                    break;
                case "--columns":
                    columns = ParsePositiveInteger(value, "columns", MaxExcelColumns);
                    break;
                case "--output":
                    output = value;
                    break;
                default:
                    throw new ArgumentException($"Unknown argument: {flag}");
            }
        }

        if (string.IsNullOrWhiteSpace(output))
        {
            throw new ArgumentException("output must not be empty");
        }

        return new Config(rows, columns, Path.GetFullPath(output));
    }

    private static int ParsePositiveInteger(string value, string name, int maximum)
    {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed)
            || parsed <= 0
            || parsed > maximum)
        {
            throw new ArgumentException($"{name} must be an integer from 1 to {maximum}");
        }
        return parsed;
    }

    private static void Run(Config config)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(config.Output)!);

        var totalStart = Stopwatch.GetTimestamp();
        var document = SpreadsheetDocument.Create(config.Output, SpreadsheetDocumentType.Workbook);
        var disposed = false;
        double writeMs;
        double saveMs;

        try
        {
            var workbookPart = document.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Data",
                    }
                )
            );

            var writeStart = Stopwatch.GetTimestamp();
            WriteWorksheet(worksheetPart, config.Rows, config.Columns);
            writeMs = Stopwatch.GetElapsedTime(writeStart).TotalMilliseconds;

            var saveStart = Stopwatch.GetTimestamp();
            workbookPart.Workbook.Save();
            disposed = true;
            document.Dispose();
            saveMs = Stopwatch.GetElapsedTime(saveStart).TotalMilliseconds;
        }
        finally
        {
            if (!disposed)
            {
                document.Dispose();
            }
        }

        var totalMs = Stopwatch.GetElapsedTime(totalStart).TotalMilliseconds;
        var process = Process.GetCurrentProcess();
        var maxRssBytes = MaxRssBytes(process);
        var outputBytes = new FileInfo(config.Output).Length;
        var result = new
        {
            implementation = "csharp/openxml-sdk",
            libraryVersion = typeof(SpreadsheetDocument).Assembly.GetName().Version?.ToString() ?? "unknown",
            runtimeVersion = RuntimeInformation.FrameworkDescription,
            mode = "streaming",
            writeMode = "cell",
            rows = config.Rows,
            columns = config.Columns,
            cells = (long)config.Rows * config.Columns,
            writeMs,
            saveMs,
            totalMs,
            maxRssBytes,
            outputBytes,
        };
        Console.WriteLine(JsonSerializer.Serialize(result));
    }

    private static long? MaxRssBytes(Process process)
    {
        process.Refresh();
        if (process.PeakWorkingSet64 > 0)
        {
            return process.PeakWorkingSet64;
        }
        if (IntPtr.Size != 8 || OperatingSystem.IsWindows())
        {
            return null;
        }

        try
        {
            if (GetResourceUsage(0, out var usage) != 0 || usage.MaxRss <= 0)
            {
                return null;
            }
            return OperatingSystem.IsMacOS()
                ? usage.MaxRss
                : checked(usage.MaxRss * 1_024);
        }
        catch (DllNotFoundException)
        {
            return null;
        }
        catch (EntryPointNotFoundException)
        {
            return null;
        }
    }

    private static void WriteWorksheet(WorksheetPart worksheetPart, int rows, int columns)
    {
        var columnNames = new string[columns];
        for (var column = 0; column < columns; column++)
        {
            columnNames[column] = ColumnName(column + 1);
        }

        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            CloseOutput = false,
        };
        using var stream = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, settings);

        writer.WriteStartDocument();
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        writer.WriteStartElement("dimension", SpreadsheetNamespace);
        writer.WriteAttributeString("ref", $"A1:{columnNames[^1]}{rows}");
        writer.WriteEndElement();
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);

        for (var row = 0; row < rows; row++)
        {
            var excelRow = (row + 1).ToString(CultureInfo.InvariantCulture);
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", excelRow);

            for (var column = 0; column < columns; column++)
            {
                writer.WriteStartElement("c", SpreadsheetNamespace);
                writer.WriteAttributeString("r", string.Concat(columnNames[column], excelRow));
                if (column % 2 == 0)
                {
                    writer.WriteAttributeString("t", "inlineStr");
                    writer.WriteStartElement("is", SpreadsheetNamespace);
                    writer.WriteElementString("t", SpreadsheetNamespace, $"Row {row} Col {column}");
                    writer.WriteEndElement();
                }
                else
                {
                    var value = (double)row * column * 0.123;
                    writer.WriteElementString(
                        "v",
                        SpreadsheetNamespace,
                        value.ToString("R", CultureInfo.InvariantCulture)
                    );
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static string ColumnName(int column)
    {
        Span<char> buffer = stackalloc char[3];
        var offset = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--offset] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(buffer[offset..]);
    }
}
