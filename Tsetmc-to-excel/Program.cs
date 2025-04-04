using static TseTmcToExcel.ExcelTools;
using static TseTmcToExcel.TseTmcTools;
using static TseTmcToExcel.IO;

namespace TseTmcToExcel
{
    public static class Program
    {
        public static async Task Main()
        {
            while (true) // Infinite loop to keep fetching data until the user exits
            {
                // Retrieve user input or default settings from configuration
                IO.Initialize();
                var counter = 0; // Counter to track the number of data fetch iterations

                Console.Clear();
                Console.WriteLine("Press 'Q' to stop the data fetching process.\n");

                // Create a cancellation token to handle user exit request
                using var cts = new CancellationTokenSource();
                _ = Task.Run(() =>
                {
                    while (Console.ReadKey(true).Key != ConsoleKey.Q) { } // Wait for 'Q' key press
                    cts.Cancel(); // Cancel the ongoing tasks
                });

                // Start continuous data fetching loop until 'Q' is pressed
                while (!cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        counter++;
                        Console.WriteLine("______________________");
                        Console.WriteLine($"#{counter} Started at {DateTime.Now.ToString("HH:mm:ss")}\n");

                        _ = UpdateData(); // Fetch and update data asynchronously
                        await Task.Delay(UpdateInterval * 1000, cts.Token); // Wait before fetching data again

                        Console.Clear();
                        Console.WriteLine("Press 'Q' to stop the data fetching process.\n");
                    }
                    catch (Exception)
                    {
                        // Exception handling is left empty
                    }
                }

                Console.WriteLine("Data fetching stopped.\n");
            }
        }

        /// <summary>
        /// Fetches data from APIs, validates it, and saves it to an Excel file.
        /// </summary>
        public static async Task UpdateData()
        {
            // Split multiple parameters and sheet names into lists for processing
            List<string> urlParamsList = ApiParameter.Replace(" ", "").Split(',').ToList();
            List<string> sheetnamesList = SheetName.Replace(" ", "").Split(',').ToList();
            long longestID = 0; // Track the highest ID found in the existing Excel sheets

            // Determine the latest ID from all sheets
            foreach (string sheet in sheetnamesList)
            {
                longestID = Math.Max(longestID, GetLatestId(ExcelFileName, sheet));
            }
            longestID++; // Increment to ensure new data has a unique ID

            // Dictionary to store retrieved data from API responses
            Dictionary<int, Dictionary<string, string>> dataCollection = new Dictionary<int, Dictionary<string, string>>();

            await Parallel.ForAsync(0, urlParamsList.Count, async (i, cancellationToken) =>
            {
                var closingPriceData = new Dictionary<string, string> { { "SharedID", longestID.ToString() } };
                bool isDataValid = true; // Flag to track data validity

                // Fetch closing price data if required
                if (SelectedItems.Any(x => ClosingItems.Contains(x)))
                {
                    var result = await GetClosingPriceInfo(urlParamsList[i], SelectedItems);
                    (isDataValid, closingPriceData) = CombineValidData(result, closingPriceData);
                }

                // Fetch ETF data if required
                if (SelectedItems.Any(x => EtfItems.Contains(x)) && isDataValid)
                {
                    // Merge ETF data into the existing data set
                    var etfData =  await GetETFByInsCode(urlParamsList[i]);
                    (isDataValid, closingPriceData) = CombineValidData(etfData, closingPriceData);
                }

                // Store valid data only
                if (isDataValid)
                    dataCollection.Add(i, closingPriceData);
            });

            // Save valid data to Excel
            if (dataCollection.Count == urlParamsList.Count)
            {
                foreach (var data in dataCollection)
                    SaveToExcel(ExcelFileName, sheetnamesList[data.Key], data.Value);
            }
            else
            {
                Console.WriteLine("No valid data received to save!");
            }
        }

        /// <summary>
        /// Combines valid data from input dictionary into result dictionary while validating the values.
        /// </summary>
        /// <param name="input">Dictionary containing input data to validate and combine</param>
        /// <param name="result">Dictionary to store valid key-value pairs from input</param>
        /// <returns>
        /// Tuple containing:
        /// - boolean indicating if all data is valid (true) or not (false)
        /// - the result dictionary with valid key-value pairs added
        /// </returns>
        private static (bool, Dictionary<string, string>) CombineValidData(Dictionary<string, string> input, Dictionary<string, string> result)
        {
            // Create a local copy of the result dictionary to avoid modifying the shared one during processing
            var localResult = new Dictionary<string, string>(result);
            var isDataValid = true;

            // Early exit if input is null or empty
            if (input == null || input.Count == 0)
            {
                return (false, result);
            }

            foreach (var item in input)
            {
                // First validation: Check if value is null or empty
                if (string.IsNullOrEmpty(item.Value))
                {
                    isDataValid = false;
                    break;
                }

                // Second validation: Check NonZeroItems (assuming it's thread-safe or immutable)
                if (NonZeroItems.Contains(item.Value) && int.TryParse(item.Value, out int value) && value == 0)
                {
                    isDataValid = false;
                    break;
                }

                // If we get here, validations passed - add to local result
                localResult[item.Key] = item.Value;
            }

            // Only update the shared result if all validations passed
            if (isDataValid)
            {
                foreach (var item in localResult)
                {
                    result[item.Key] = item.Value;
                }
            }

            return (isDataValid, result);
        }
    }
}