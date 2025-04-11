using static TseTmcToExcel.ExcelTools;
using static TseTmcToExcel.TseTmcTools;
using static TseTmcToExcel.IO;
using OfficeOpenXml;
using static System.Runtime.InteropServices.JavaScript.JSType;

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

                var userInput = "";
                while (!userInput!.Equals("y"))
                {
                    Console.Clear();
                    Console.Write($"Do you want to start again? Y/N");
                    userInput = Console.ReadKey().KeyChar.ToString().ToLower();
                    if (userInput.Equals("n"))
                        Environment.Exit(0);
                }
            }
        }

        /// <summary>
        /// Fetches data from APIs, validates it, and saves it to an Excel file.
        /// </summary>
        public static async Task UpdateData()
        {
            // Split multiple parameters and sheet names into lists for processing
            List<string> sheetnamesList = new List<string>()
            {
                OverviewSheetName, StockSheetName
            };

            // Track the highest ID found in the existing Excel sheets
            long longestID = 0;

            // Determine the latest ID from all sheets
            foreach (string sheet in sheetnamesList)
            {
                longestID = Math.Max(longestID, GetLatestId(ExcelFileName, sheet));
            }
            longestID++; // Increment to ensure new data has a unique ID

            // Dictionary to store retrieved data from API responses
            Dictionary<int, Dictionary<string, string>> dataCollection = new Dictionary<int, Dictionary<string, string>>();

            // Fetch data for each url param parallel or linear
            if (isParallel)
            {
                await Parallel.ForAsync(0, ApiParameterList.Count, async (i, cancellationToken) =>
                {
                    dataCollection = await FetchValidData(dataCollection, longestID, i);
                });
            }
            else
            {
                for (int i = 0; i < ApiParameterList.Count; i++)
                {
                    dataCollection = await FetchValidData(dataCollection, longestID, i);
                }
            }

            // Open the excel file
            ExcelPackage package = new ExcelPackage();
            FileInfo file = new FileInfo(ExcelFileName);
            var changes = new List<string>();
            var excelFile = OpenExcel(package, file);

            // Save valid data to Excel
            if (dataCollection.Count == ApiParameterList.Count)
                foreach (var data in dataCollection)
                {
                    package = AddToExcel(excelFile, StockSheetName, data.Value);
                    
                    changes.Add(StockNameList[data.Key]);
                }
            else
                Console.WriteLine("No valid data received to save!");


            if (SelectedItems.Any(x => MarketOverviewItems.Contains(x)))
            {
                // Validate and Add GeneralStockData to the closingPriceData
                var marketOverview = new Dictionary<string, string>{ { "SharedID", longestID.ToString() } };
                marketOverview = CombineValidData(await GetMarketOverview(), marketOverview).Item2;

                // Save OverView data to Excel
                if (marketOverview != null && marketOverview.Any())
                {
                    package = AddToExcel(excelFile, OverviewSheetName, marketOverview);
                    changes.Add(OverviewSheetName);
                }
                else
                {
                    Console.WriteLine("No valid data received to save!");
                }
            }
            // save the changed file
            SaveExcel(package, file, changes);
        }

        /// <summary>
        /// Fetch data from TSETMC and put if its valid add to a data collection
        /// </summary>
        /// <param name="inputCollection">the input data collection</param>
        /// <param name="longestID">the longest shared ID</param>
        /// <param name="urlParam">the url param need to be fetched</param>
        /// <param name="i">the index of the stocks list</param>
        /// <returns>the modified data collection that result has been added into</returns>
        private static async Task<Dictionary<int, Dictionary<string,string>>> FetchValidData(Dictionary<int, Dictionary<string, string>> inputCollection, long longestID, int i)
        {
            var dataCollection = inputCollection;
            var closingPriceData = new Dictionary<string, string> { { "SharedID", longestID.ToString() }, { "Stock", StockNameList[i] } };
            bool isDataValid = true; // Flag to track data validity

            // Fetch closing price data if required
            if (SelectedItems.Any(x => ClosingItems.Contains(x)))
            {
                var result = await GetClosingPriceInfo(ApiParameterList[i], SelectedItems);
                (isDataValid, closingPriceData) = CombineValidData(result, closingPriceData);
            }

            // Fetch ETF data if required
            if (SelectedItems.Any(x => EtfItems.Contains(x)) && isDataValid)
            {
                // Merge ETF data into the existing data set
                var etfData = await GetETFByInsCode(ApiParameterList[i]);
                (isDataValid, closingPriceData) = CombineValidData(etfData, closingPriceData);
            }

            // Store valid data only
            if (isDataValid)
                dataCollection.Add(i, closingPriceData);

            return dataCollection;
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
                if (string.IsNullOrWhiteSpace(item.Value))
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