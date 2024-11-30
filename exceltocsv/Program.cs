using System;
using System.IO;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Collections.Generic;
using System.Text.RegularExpressions;

class Program
{
    enum ColumnType
    {
        Text,
        Numeric,
        Date
    }

    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Console.WriteLine("Enter Excel file path:");
        string excelPath = Console.ReadLine();

        Console.WriteLine("Enter output CSV file path:");
        string csvPath = Console.ReadLine();

        Console.WriteLine("Enter worksheet name (or press Enter for first worksheet):");
        string worksheetName = Console.ReadLine();

        Console.WriteLine("Enter date format (e.g., 'yyyy-MM-dd' or press Enter for default):");
        string dateFormat = Console.ReadLine();

        try
        {
            ConvertExcelToCsv(excelPath, csvPath, worksheetName, dateFormat);
            Console.WriteLine("Conversion completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static void ConvertExcelToCsv(string excelPath, string csvPath, string worksheetName = "", string dateFormat = "")
    {
        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            ExcelWorksheet worksheet;
            if (string.IsNullOrEmpty(worksheetName))
            {
                worksheet = package.Workbook.Worksheets[0];
            }
            else
            {
                worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet == null)
                {
                    var availableWorksheets = string.Join(", ", 
                        package.Workbook.Worksheets.Select(w => w.Name));
                    throw new Exception(
                        $"Worksheet '{worksheetName}' not found. Available worksheets: {availableWorksheets}");
                }
            }

            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            // Determine column types from the first 25 rows or all rows if less than 25
            var columnTypes = DetermineColumnTypes(worksheet, columns, Math.Min(rows, 26));

            using (var sw = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                // Write headers without quotes
                for (int col = 1; col <= columns; col++)
                {
                    string value = worksheet.Cells[1, col].Text ?? "";
                    sw.Write(value);
                    if (col < columns)
                    {
                        sw.Write(",");
                    }
                }
                sw.WriteLine();

                // Write data rows
                for (int row = 2; row <= rows; row++)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        string value;

                        if (cell.Value == null)
                        {
                            value = "";
                        }
                        else if (cell.Value is DateTime date)
                        {
                            value = !string.IsNullOrEmpty(dateFormat) 
                                ? date.ToString(dateFormat) 
                                : date.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            value = cell.Text; // Using .Text instead of .ToString() to get formatted value
                        }

                        // Apply quoting based on column type
                        if (columnTypes[col - 1] == ColumnType.Text || HasLeadingZeros(value))
                        {
                            sw.Write($"\"{value.Replace("\"", "\"\"")}\"");
                        }
                        else if (columnTypes[col - 1] == ColumnType.Date)
                        {
                            sw.Write($"\"{value}\"");
                        }
                        else
                        {
                            sw.Write(value);
                        }

                        if (col < columns)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.WriteLine();
                }
            }
        }
    }

    static List<ColumnType> DetermineColumnTypes(ExcelWorksheet worksheet, int columns, int maxRows)
    {
        var columnTypes = new List<ColumnType>();

        for (int col = 1; col <= columns; col++)
        {
            bool hasNonNumeric = false;
            bool hasAnyValue = false;
            bool allDates = true;
            bool hasLeadingZeros = false;

            // Check first 25 rows (or less if worksheet is smaller)
            for (int row = 2; row < maxRows; row++)
            {
                var cell = worksheet.Cells[row, col];
                string cellText = cell.Text; // Get formatted value

                if (cell.Value != null)
                {
                    hasAnyValue = true;

                    // Check for leading zeros
                    if (HasLeadingZeros(cellText))
                    {
                        hasLeadingZeros = true;
                        break;
                    }

                    if (cell.Value is DateTime)
                    {
                        continue; // Keep allDates true
                    }
                    else
                    {
                        allDates = false;
                    }

                    if (!(cell.Value is double || cell.Value is decimal || cell.Value is float || cell.Value is int || cell.Value is long))
                    {
                        hasNonNumeric = true;
                        break;
                    }
                }
            }

            // Determine column type
            if (!hasAnyValue || hasNonNumeric || hasLeadingZeros)
            {
                columnTypes.Add(ColumnType.Text);
            }
            else if (allDates)
            {
                columnTypes.Add(ColumnType.Date);
            }
            else
            {
                columnTypes.Add(ColumnType.Numeric);
            }

            Console.WriteLine($"Column {col}: {columnTypes[col-1]}"); // Debug information
        }

        return columnTypes;
    }

    static bool HasLeadingZeros(string value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        
        // Check if the value starts with 0 and has more than one character
        if (value.StartsWith("0") && value.Length > 1)
        {
            // Make sure it's not a decimal number starting with 0.
            if (!value.StartsWith("0."))
            {
                return true;
            }
        }
        return false;
    }
}