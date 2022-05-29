/*
 * The MIT License (MIT)
 *
 * Copyright (c) 2015 Roy Xu
 *
 * @since 12/31/2015 18:53:05
 *
 * @author Roy Xu
 *
*/

using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

namespace kitapp
{
    class Program
    {
        private const String DEFAULT_OUTPUT_PATH = "out.xlsx";
        private const String DEFAULT_EXPORT_TYPE = "--xlsx";

        static void Main()
        {
            //Console.ReadKey(); // for debug with attach to process .. [Ctrl+Alt+P]

            String[] commandLineArgs = Environment.GetCommandLineArgs();
            String command = Path.GetFileNameWithoutExtension(commandLineArgs[0]);

            // display this help text and exit
            if (commandLineArgs.Length == 1 ||
                (commandLineArgs.Length == 2 && "--help".Equals(commandLineArgs[1])))
            {
                DisplayHelpText(command);
                return;
            }

            // display version information and exit
            if (commandLineArgs.Length == 2 && "--version".Equals(commandLineArgs[1]))
            {
                DisplayVersionInformation();
                return;
            }

            // export json text to *.xlsx or *.docx
            try
            {
                ExportToOffice(commandLineArgs);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
                DisplayHelpText(command);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandLineArgs"></param>
        /// <exception cref="System.IO.FileNotFoundException">Invalid option: the file specified in option '--file=&lt;file&gt;' was not found.</exception>
        private static void ExportToOffice(String[] commandLineArgs)
        {
            var command = Path.GetFileName(commandLineArgs[0]);
            var json = GetJsonTextFromCommandLineArgs(commandLineArgs);
            if (String.IsNullOrEmpty(json))
            {
                DisplayHelpText(command);
                return;
            }
            var outputPath = GetOutputPathFromCommandLineArgs(commandLineArgs) ?? DEFAULT_OUTPUT_PATH;
            var exportType = GetExportTypeFromCommandLineArgs(commandLineArgs) ?? DEFAULT_EXPORT_TYPE;

            switch (exportType)
            {
                case "--xlsx":
                    ExportToXlsx(json, outputPath);
                    break;
                case "--docx":
                    ExportToDocx(json, outputPath);
                    break;
                default:
                    DisplayHelpText(command);
                    break;
            }
        }

        private static string? GetExportTypeFromCommandLineArgs(string[] commandLineArgs)
        {
            string? exportType = null;
            foreach (var arg in commandLineArgs)
            {
                switch (arg)
                {
                    case "--xlsx":
                    case "--docx":
                        exportType = arg;
                        break;
                }
            }
            return exportType;
        }

        private static string? GetOutputPathFromCommandLineArgs(String[] commandLineArgs)
        {
            string? outputPath = null;
            foreach (var arg in commandLineArgs)
            {
                if (arg.StartsWith("--out"))
                {
                    outputPath = arg.Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
                    break;
                }
            }
            return outputPath;
        }

        private static string? GetJsonTextFromCommandLineArgs(String[] commandLineArgs)
        {
            string? json = null;
            if (commandLineArgs.Length == 2 && !commandLineArgs[1].StartsWith("--"))
            { // the first command argument as json text if only one argument provided.
                json = commandLineArgs[1];
            }
            else
            {
                foreach (var arg in commandLineArgs)
                {
                    if (arg.StartsWith("--json"))
                    {
                        json = arg.Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
                        break;
                    }
                    if (arg.StartsWith("--file"))
                    {
                        var path = arg.Split(new[] { '=' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
                        if (File.Exists(path))
                        {
                            json = File.ReadAllText(path);
                        }
                        else
                        {
                            throw new FileNotFoundException(String.Format("Invalid option: the file specified in option '{0}' was not found.", arg, CultureInfo.InvariantCulture));
                        }
                        break;
                    }
                }
            }
            return json;
        }

        private static void ExportToXlsx(String json, [DisallowNull] String outputPath)
        {
            outputPath = AmendOutputPath(outputPath, ".xlsx");
            using var o = JsonDocument.Parse(json);
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("sheet1");
            int rownum = 0;
            foreach (var token in o.RootElement.GetProperty("data").EnumerateArray())
            {
                var row = sheet.CreateRow(rownum++);
                var column = 0;
                foreach (var text in token.EnumerateObject())
                {
                    row.CreateCell(column++).SetCellValue(text.Value.GetString());
                }
            }
            using (var fs = File.Create(outputPath))
            {
                workbook.Write(fs);
            }
        }

        private static void ExportToDocx(string json, [DisallowNull] string outputPath)
        {
            outputPath = AmendOutputPath(outputPath, ".docx");
            using var jobject = JsonDocument.Parse(json);
            var data = jobject.RootElement.GetProperty("data");
            var rows = data.EnumerateArray().Count();
            var cols = 0;
            if (rows > 0)
            {
                cols = data.EnumerateArray().First().EnumerateObject().Count();
            }
            var document = new XWPFDocument();
            var table = document.CreateTable(rows, cols);
            //for (int col = 0; col < cols; col++) { // not work
            //    table.SetColumnWidth(col, 100);
            //}

            var row = 0;
            foreach (var token in data.EnumerateArray())
            {
                int pos = 0;
                foreach (var text in token.EnumerateObject())
                {
                    table.GetRow(row).GetCell(pos++).SetText(text.Value.GetString());
                }
                row++;
            }
            using (var fs = File.Create(outputPath))
            {
                document.Write(fs);
            }
        }

        private static string AmendOutputPath(string outputPath, string extension)
        {
            var location = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(location))
            {
                Directory.CreateDirectory(location!);
            }
            outputPath = Path.Combine(location!, Path.GetFileNameWithoutExtension(outputPath) + extension);
            return outputPath;
        }

        private static void DisplayVersionInformation()
        {
            Console.WriteLine(__Version.Version);
        }

        private static void DisplayHelpText(String command)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("{0} [--help] [--version] [--file=<file>] [--json=<a character sequence>] [-out=<file>] [--xlsx] [--docx] [<a character sequence>]", command);
            Console.WriteLine("Options:");
            Console.WriteLine("{0,-20}display this help text and exit.", "--help");
            Console.WriteLine("{0,-20}display version information and exit.", "--version");
            Console.WriteLine("{0,-20}select a json file.", "--file=<file>");
            Console.WriteLine("{0,-20}", "--json=<a character sequence>");
            Console.WriteLine("{0,-20}select a character sequence that represents a json object.", String.Empty);
            Console.WriteLine("{0,-20}select export type, default: --xlsx.", "--xlsx,--docx");
            Console.WriteLine("{0,-20}select output file path, default: out.xlsx or out.docx.", "--out=<file>");
            Console.WriteLine("{0,-20}", "<a character sequence>");
            Console.WriteLine("{0,-20}select a character sequence that represents a json object, only if one positional parameter (not start with '--') provided, Jsonkit will use it as json text and export a *.xlsx.", String.Empty);
            Console.WriteLine("Examples:");
            String sampleJson = @"{""data"":[{""symbol"":""601288"",""name"":""农业银行"",""value"":""13.05""},{""symbol"":""600291"",""name"":""西水股份"",""value"":""12.69""}]}";
            Console.WriteLine("{0,10} --file=sample.json --xlsx --out=out.xlsx", command);
            Console.WriteLine("{0,10} --file=sample.json --xlsx", command);
            Console.WriteLine("{0,10} --file=sample.json", command);
            Console.WriteLine("{0,10} --json={1} --xlsx --out=out.xlsx", command, sampleJson);
            Console.WriteLine("{0,10} --json={1} --xlsx", command, sampleJson);
            Console.WriteLine("{0,10} --json={1}", command, sampleJson);
            Console.WriteLine("{0,10} {1}", command, sampleJson);
        }
    }
}
