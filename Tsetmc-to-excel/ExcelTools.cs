using System.Globalization;
using OfficeOpenXml;

namespace TseTmcToExcel
{
    public static class ExcelTools
    {
        /// <summary>
        /// Saves data to an Excel file, creating missing columns if necessary.
        /// </summary>
        /// <param name="filePath">The file path of the Excel file.</param>
        /// <param name="sheetName">The name of the sheet to write data to.</param>
        /// <param name="data">A dictionary containing key-value pairs to store in the sheet.</param>
        public static void SaveToExcel(string filePath, string sheetName, Dictionary<string, string> data)
        {
            try
            {
                FileInfo file = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context for EPPlus

                // Open existing Excel file if it exists, otherwise create a new package
                using var package = file.Exists ? new ExcelPackage(file) : new ExcelPackage();

                // Get the worksheet by name or create a new one if it doesn't exist
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
                                ?? package.Workbook.Worksheets.Add(sheetName);

                // Store existing column headers in a list
                List<string> existingHeaders = new List<string>();
                if (worksheet.Dimension != null)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        existingHeaders.Add(worksheet.Cells[1, col].Text);
                    }
                }

                // Ensure all keys in the data dictionary exist as column headers
                int colIndex = 1;
                foreach (var key in data.Keys)
                {
                    if (!existingHeaders.Contains(key))
                    {
                        worksheet.Cells[1, existingHeaders.Count + 1].Value = key; // Add new header
                        existingHeaders.Add(key);
                    }
                    colIndex++;
                }

                // Determine the next available row for inserting new data
                int newRow = worksheet.Dimension?.Rows + 1 ?? 2;
                colIndex = 1;

                // Insert data into the appropriate columns
                foreach (var key in existingHeaders)
                {
                    if (data.ContainsKey(key))
                    {
                        worksheet.Cells[newRow, colIndex].Value = ParseString(data[key]);
                    }
                    colIndex++;
                }

                // Save the Excel file
                package.SaveAs(file);
                Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}\t{sheetName}");
            }
            catch (Exception ex)
            {
                // Handle and log errors during saving
                Console.WriteLine($"Error saving Excel: {ex.Message}");
                ex = ex.InnerException!;
                while (ex != null)
                {
                    Console.WriteLine($"InnerException: {ex.Message}");
                    ex = ex.InnerException!;
                }
            }
        }

        /// <summary>
        /// Retrieves the latest ID from a specific column in an Excel sheet.
        /// </summary>
        /// <param name="filePath">The file path of the Excel file.</param>
        /// <param name="sheetName">The name of the sheet to read from.</param>
        /// <param name="columnIndex">The column index where IDs are stored (default is 1).</param>
        /// <returns>The highest numerical ID found in the specified column.</returns>
        public static long GetLatestId(string filePath, string sheetName, int columnIndex = 1)
        {
            try
            {
                FileInfo file = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context

                // Open the existing Excel file
                using var package = new ExcelPackage(file);
                var worksheet = package.Workbook.Worksheets[sheetName];

                // If the sheet is empty or does not exist, return 0
                if (worksheet == null || worksheet.Dimension == null)
                    return 0;

                int lastRow = worksheet.Dimension.Rows;
                long latestId = 0;

                // Iterate through rows (assuming row 1 is a header)
                for (int row = 2; row <= lastRow; row++)
                {
                    var cellValue = worksheet.Cells[row, columnIndex].Value;
                    if (cellValue != null && long.TryParse(cellValue.ToString(), out long id))
                    {
                        latestId = Math.Max(latestId, id); // Track the highest ID found
                    }
                }

                return latestId;
            }
            catch (Exception ex)
            {
                // Handle errors during reading
                Console.WriteLine($"Error retrieving latest ID: {ex.Message}");
                return 0;
            }
        }


        /// <summary>
        /// Attempts to parse the input string as a number.
        /// If successful, returns the numeric value (as a double).
        /// If not, returns the original string.
        /// </summary>
        /// <param name="input">The input string to parse.</param>
        /// <returns>A double if the input is numeric; otherwise, the original string.</returns>
        public static object ParseString(string input)
        {
            // Try to parse the input string to a double using invariant culture (for consistent decimal format)
            if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
            {
                // If parsing succeeds, return the numeric value
                return number;
            }

            // If parsing fails, return the original string
            return input;
        }
    }

}
