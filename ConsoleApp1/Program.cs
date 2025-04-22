using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

class Program
{
    static void Main()
    {
        // Set the EPPlus license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Get file paths and version input from user
        Console.Write("Enter OLD Excel file path: ");
        string oldPath = Console.ReadLine();

        Console.Write("Enter NEW Excel file path: ");
        string newPath = Console.ReadLine();

        Console.Write("Enter OUTPUT Excel file path: ");
        string outputPath = Console.ReadLine();

        Console.Write("Enter version label (e.g., v1.4.00000): ");
        string versionLabel = Console.ReadLine()?.Trim();
        string updatedColumnName = $"Updated in {versionLabel}";

        Console.Write("Enter numeric threshold to determine if fix is required (e.g., 10): ");
        int.TryParse(Console.ReadLine(), out int threshold);

        // Read both Excel files into memory
        var oldData = ReadExcelToDictionary(oldPath);
        var newData = ReadExcelToDictionary(newPath);

        var headers = newData.FirstOrDefault()?.Keys.ToList() ?? new List<string>();

        Console.Write("Do you want to compare OLD and NEW files? (y/n): ");
        if (Console.ReadLine()?.Trim().ToLower() == "y")
        {
            // Collect key columns and columns to compare from user
            Console.Write("Enter KEY column letters (comma-separated, e.g., A,B): ");
            var keyLetters = Console.ReadLine().Split(',').Select(c => c.Trim()).ToList();
            var keyColumns = keyLetters.Select(l => GetHeaderNameFromLetter(l, headers)).Where(h => h != null).ToList();

            Console.Write("Enter columns to compare (e.g., G): ");
            var valLetters = Console.ReadLine().Split(',').Select(c => c.Trim()).ToList();
            var valColumns = valLetters.Select(l => GetHeaderNameFromLetter(l, headers)).Where(h => h != null).ToList();

            // Create a lookup for old data based on keys
            var oldDict = BuildDictionary(oldData, keyColumns);

            // Compare and add new columns to the new data
            foreach (var row in newData)
            {
                string key = string.Join("|", keyColumns.Select(k => row.ContainsKey(k) ? row[k] : ""));
                string result = "", updatedIn = "No", fixRequired = "No", isFixed = "No";

                if (oldDict.TryGetValue(key, out var oldRow))
                {
                    foreach (var col in valColumns)
                    {
                        string oldVal = oldRow.ContainsKey(col) ? oldRow[col] : "";
                        string newVal = row.ContainsKey(col) ? row[col] : "";

                        if (oldVal != newVal)
                        {
                            result += $"Before: {oldVal}\nAfter: {newVal}\n\n";
                            updatedIn = "Yes";

                            if (int.TryParse(newVal, out int newValue) && newValue > threshold)
                            {
                                fixRequired = "Yes";
                            }
                        }
                    }
                }
                else
                {
                    updatedIn = "Yes";
                    foreach (var col in valColumns)
                    {
                        if (row.ContainsKey(col) && int.TryParse(row[col], out int newValue) && newValue > threshold)
                        {
                            fixRequired = "Yes";
                            break;
                        }
                    }
                }

                row[updatedColumnName] = updatedIn;
                row["Fix Required"] = fixRequired;
                row["Fixed?"] = isFixed;
                row["Note"] = result.Trim();
            }
        }

        // Create output Excel file
        var outputPkg = new ExcelPackage();

        // Write complete data to "AllData" sheet
        var allSheet = outputPkg.Workbook.Worksheets.Add("AllData");
        WriteTable(allSheet, newData);

        // Optionally divide into multiple sheets
        Console.Write("Do you want to divide the NEW file into multiple sheets? (y/n): ");
        if (Console.ReadLine()?.Trim().ToLower() == "y")
        {
            Console.Write("Enter column letter to divide by (e.g., 'B'): ");
            string splitColLetter = Console.ReadLine().Trim();
            string splitColumn = GetHeaderNameFromLetter(splitColLetter, headers);

            if (!string.IsNullOrEmpty(splitColumn))
            {
                var grouped = newData.GroupBy(row => row.ContainsKey(splitColumn) ? row[splitColumn] : "Empty");

                foreach (var group in grouped)
                {
                    string name = string.IsNullOrEmpty(group.Key) ? "Empty" : group.Key;
                    string safeName = name.Length > 31 ? name.Substring(0, 31) : name;
                    var sheet = outputPkg.Workbook.Worksheets.Add(safeName);
                    WriteTable(sheet, group.ToList());
                }
            }
        }

        outputPkg.SaveAs(new FileInfo(outputPath));
        Console.WriteLine("✅ Operation completed. Output saved.");
    }

    // Reads Excel and returns list of dictionaries (rows)
    static List<Dictionary<string, string>> ReadExcelToDictionary(string path)
    {
        var pkg = new ExcelPackage(new FileInfo(path));
        var ws = pkg.Workbook.Worksheets[0];

        var headers = new List<string>();
        for (int col = 1; col <= ws.Dimension.End.Column; col++)
        {
            headers.Add(ws.Cells[1, col].Text);
        }

        var rows = new List<Dictionary<string, string>>();
        for (int row = 2; row <= ws.Dimension.End.Row; row++)
        {
            var dict = new Dictionary<string, string>();
            for (int col = 1; col <= headers.Count; col++)
            {
                dict[headers[col - 1]] = ws.Cells[row, col].Text;
            }
            rows.Add(dict);
        }

        return rows;
    }

    // Builds dictionary using keys for lookup
    static Dictionary<string, Dictionary<string, string>> BuildDictionary(List<Dictionary<string, string>> data, List<string> keys)
    {
        var dict = new Dictionary<string, Dictionary<string, string>>();
        foreach (var row in data)
        {
            var key = string.Join("|", keys.Select(k => row.ContainsKey(k) ? row[k] : ""));
            dict[key] = row;
        }
        return dict;
    }

    // Writes a list of dictionaries (table) to worksheet
    static void WriteTable(ExcelWorksheet sheet, List<Dictionary<string, string>> data)
    {
        if (data.Count == 0) return;

        var headers = data[0].Keys.ToList();
        for (int col = 0; col < headers.Count; col++)
        {
            sheet.Cells[1, col + 1].Value = headers[col];
        }

        for (int row = 0; row < data.Count; row++)
        {
            for (int col = 0; col < headers.Count; col++)
            {
                sheet.Cells[row + 2, col + 1].Value = data[row].ContainsKey(headers[col]) ? data[row][headers[col]] : "";
            }
        }
    }

    // Gets column header name from column letter (A, B, C, ...)
    static string GetHeaderNameFromLetter(string letter, List<string> headers)
    {
        int colNum = ColumnLetterToNumber(letter);
        if (colNum <= headers.Count)
            return headers[colNum - 1];
        return null;
    }

    // Converts column letter to index (e.g., A => 1, B => 2, AA => 27)
    private static int ColumnLetterToNumber(string columnLetter)
    {
        int sum = 0;
        foreach (char c in columnLetter.ToUpper())
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
    }
}