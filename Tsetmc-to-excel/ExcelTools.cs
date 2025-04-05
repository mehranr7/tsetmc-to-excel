using System.Globalization;
using OfficeOpenXml;

namespace TseTmcToExcel
{
    public static class ExcelTools
    {
        /// <summary>
        /// Opens an existing Excel file or creates a new Excel package if the file doesn't exist
        /// </summary>
        /// <param name="package">ExcelPackage object (can be null)</param>
        /// <param name="file">FileInfo object representing the Excel file</param>
        /// <returns>Initialized ExcelPackage object</returns>
        public static ExcelPackage OpenExcel(ExcelPackage package, FileInfo file)
        {
            try
            {
                // Set license context for EPPlus to NonCommercial (required for EPPlus 5+)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Open existing Excel file if it exists, otherwise create a new package
                package = file.Exists ? new ExcelPackage(file) : new ExcelPackage();
            }
            catch (Exception ex)
            {
                // Handle and log errors during file opening/creation
                Console.WriteLine($"Error saving Excel: {ex.Message}");

                // Loop through and log all inner exceptions for detailed error information
                ex = ex.InnerException!;
                while (ex != null)
                {
                    Console.WriteLine($"InnerException: {ex.Message}");
                    ex = ex.InnerException!;
                }
            }

            // Return the initialized Excel package (or null if an error occurred)
            return package;
        }

        /// <summary>
        /// Saves the Excel package to a file and logs any changes made
        /// </summary>
        /// <param name="package">ExcelPackage object to save</param>
        /// <param name="file">Target file to save to</param>
        /// <param name="changes">List of changes to log</param>
        public static void SaveExcel(ExcelPackage package, FileInfo file, List<string> changes)
        {
            try
            {
                // Save the Excel file to the specified location
                package.SaveAs(file);
            }
            catch (Exception ex)
            {
                // Handle and log errors during saving
                Console.WriteLine($"Error saving Excel: {ex.Message}");

                // Loop through and log all inner exceptions for detailed error information
                ex = ex.InnerException!;
                while (ex != null)
                {
                    Console.WriteLine($"InnerException: {ex.Message}");
                    ex = ex.InnerException!;
                }
            }

            // Log the timestamp and all applied changes
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}");
            foreach (string change in changes)
                Console.WriteLine($" - {change}");
        }

        /// <summary>
        /// Adds data to an Excel file, creating missing columns if necessary.
        /// </summary>
        /// <param name="package">The Excel file.</param>
        /// <param name="sheetName">The name of the sheet to write data to.</param>
        /// <param name="data">A dictionary containing key-value pairs to store in the sheet.</param>
        public static ExcelPackage AddToExcel(ExcelPackage package, string sheetName, Dictionary<string, string> data)
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                // Handle and log errors during saving
                Console.WriteLine($"Error adding to Excel: {ex.Message}");
                ex = ex.InnerException!;
                while (ex != null)
                {
                    Console.WriteLine($"InnerException: {ex.Message}");
                    ex = ex.InnerException!;
                }
            }

            return package;
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
