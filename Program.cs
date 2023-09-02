using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Spectre.Console;
using Path = System.IO.Path;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

const int maxExecutingTasks = 6;
var timer = new Stopwatch();
var semaphore = new SemaphoreSlim(maxExecutingTasks);

var templatePath = AnsiConsole.Prompt(AskTemplatePath());
var spreadsheetPath = AnsiConsole.Prompt(AskSpreadsheetPath());
var destinationPath = AnsiConsole.Prompt(AskDestinationPath());
var shouldConvertToPdf = AnsiConsole.Confirm("Convert to pdf? (requires libreoffice)", false);

AnsiConsole.MarkupLine("Getting [green]ready[/]... ");
timer.Start();
var template = WordprocessingDocument.Open(templatePath, false);
var placeholdersTextNodes = GetPlaceholdersTextNodes(template);
template.Dispose();
var placeholders = placeholdersTextNodes.Select(x => x.Match.Groups[1].Value);
var spreadsheet = SpreadsheetDocument.Open(spreadsheetPath, false);
var rows = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>();
var header = rows.First();
var data = rows.Skip(1).Where(row => row.Elements<Cell>().Any(cell => cell.DataType != null));
var headerNamesIndexes = GetHeaderColumnsIndexes(spreadsheet, header, placeholders);
await AnsiConsole.Progress()
    .StartAsync(async ctx =>
    {
        var progressTask = ctx.AddTask("[teal]Templating[/]");
        var tasks = data.Select(async row =>
        {
            await semaphore.WaitAsync();
            var firstColumnValue = spreadsheet.GetCellValue(row, 0);
            try
            {
                var templatedDoc = GenerateTemplatedDoc(row);
                if (shouldConvertToPdf)
                {
                    await ConvertToPdf(templatedDoc, destinationPath);
                    File.Delete(templatedDoc);
                }
                else
                    File.Move(templatedDoc, Path.Join(destinationPath, Path.GetFileName(templatedDoc)));
                AnsiConsole.MarkupLine($"Templated doc [aqua]{firstColumnValue}[/] generated [green]successfully[/]!");
            }
            catch
            {
                AnsiConsole.MarkupLine($"[red]Error[/] when generating templated doc [silver]{firstColumnValue}[/]");
                // TODO: log exception to file
            }
            finally
            {
                progressTask.Increment(1);
                semaphore.Release();
            }
        }).ToList();
        progressTask.MaxValue = tasks.Count;
        await Task.WhenAll(tasks);
    });
timer.Stop();
AnsiConsole.MarkupLine($"Finished in [green]{timer.ElapsedMilliseconds / 1000}s[/].");
AnsiConsole.MarkupLine("Press any key to [green]exit[/].");
Console.ReadKey();

string GenerateTemplatedDoc(Row row)
{
    var firstColumnValue = spreadsheet.GetCellValue(row, 0);
    var (templatedDocPath, templatedDoc) = CreateTempDoc(templatePath, firstColumnValue);
    using (templatedDoc)
    {
        foreach (var (textNode, placeholder) in GetPlaceholdersTextNodes(templatedDoc))
        {
            textNode.Text = textNode.Text.Replace(
                placeholder.Value,
                spreadsheet.GetCellValue(row, headerNamesIndexes[placeholder.Groups[1].Value]));
        }
        return templatedDocPath;
    }
}

async Task ConvertToPdf(string wordPath, string pdfPath)
{
    var libreofficeBin = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
        ? @"""C:\Program Files\LibreOffice\program\soffice.com"""
        : "soffice";
    var workDir = Path.Join(Path.GetTempPath(), $"LibO_Process_{Guid.NewGuid():N}");
    var workDirUri = new Uri(workDir).AbsoluteUri;
    var processInfo = new ProcessStartInfo
    {
        FileName = libreofficeBin,
        Arguments = $@"--headless ""-env:UserInstallation={workDirUri}"" --convert-to pdf:writer_pdf_Export --outdir ""{pdfPath}"" ""{wordPath}""",
        UseShellExecute = false,
        CreateNoWindow = true,
        RedirectStandardError = true,
        RedirectStandardOutput = true
    };
    var process = new Process
    {
        StartInfo = processInfo,
        EnableRaisingEvents = true
    };
    process.Start();
    await process.WaitForExitAsync();
    Directory.Delete(workDir, true);
    if (process.ExitCode != 0) throw new Exception($"PDF conversion error, libreoffice returned exitcode {process.ExitCode}");
}

static Dictionary<string, int> GetHeaderColumnsIndexes(SpreadsheetDocument spreadsheet, Row header,
    IEnumerable<string> placeholders)
{
    return placeholders.ToDictionary(x => x,
        x => GetHeaderColumnIndex(spreadsheet, header, x) ??
             throw new Exception($"Column {x} not found in spreadsheet."));
    
    int? GetHeaderColumnIndex(SpreadsheetDocument spreadsheet, Row header, string name)
    {
        return header
            .Select((_, index) => index)
            .FirstOrDefault(x =>
                spreadsheet.GetCellValue(header, x).Equals(name, StringComparison.InvariantCultureIgnoreCase));
    }
}

static IEnumerable<(Text TextNode, Match Match)> GetPlaceholdersTextNodes(WordprocessingDocument word)
{
    return word.MainDocumentPart!.Document.Body!
        .Descendants<Text>()
        .Select(x => new { TextNode = x, PlaceHolders = Regex.Matches(x.Text, @"#(\w*?)#") })
        .Where(x => x.PlaceHolders.Count > 0)
        .SelectMany(x => x.PlaceHolders.Select(v => (x.TextNode, v)));
}

static (string Path, WordprocessingDocument word) CreateTempDoc(string templatePath, string name)
{
    var tempFilePath = GetTempFilePathWithExtension(name, "docx");
    File.Copy(templatePath, tempFilePath, true);
    return (tempFilePath, WordprocessingDocument.Open(tempFilePath, true));
}

static string GetTempFilePathWithExtension(string name, string extension)
{
    var path = Path.GetTempPath();
    var fileName = Path.ChangeExtension(name, extension);
    return Path.Combine(path, fileName);
}

static TextPrompt<string> AskTemplatePath()
{
    return new TextPrompt<string>("[green]Template[/] path: ")
        .ValidationErrorMessage("[red]Invalid.[/]")
        .Validate(path =>
        {
            // Just some basic validation
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || Path.GetExtension(path) != ".docx")
                return ValidationResult.Error();
            return ValidationResult.Success();
        });
}

static TextPrompt<string> AskSpreadsheetPath()
{
    return new TextPrompt<string>("[green]Spreadsheet[/] path: ")
        .ValidationErrorMessage("[red]Invalid[/]")
        .Validate(path =>
        {
            // Just some basic validation
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || Path.GetExtension(path) != ".xlsx")
                return ValidationResult.Error();
            return ValidationResult.Success();
        });
}

static TextPrompt<string> AskDestinationPath()
{
    return new TextPrompt<string>("[green]Destination folder[/] path: ")
        .ValidationErrorMessage("[red]Invalid[/]")
        .Validate(path =>
        {
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
                return ValidationResult.Error();
            return ValidationResult.Success();
        });
}

public static class SpreadsheetUtils
{
    public static string GetCellValue(this SpreadsheetDocument spreadsheet, Row row, int columnIndex)
    {
        var cell = row.Descendants<Cell>().ElementAt(columnIndex);
        return cell.DataType?.Value switch
        {
            CellValues.Number => cell.InnerText,
            CellValues.SharedString => spreadsheet.WorkbookPart!.GetPartsOfType<SharedStringTablePart>()
                .First()
                .SharedStringTable.ElementAt(int.Parse(cell.InnerText))
                .InnerText,
            _ => ""
        };
    }
}