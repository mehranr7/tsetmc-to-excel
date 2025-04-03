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

            // Semaphore to prevent concurrent access issues
            SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);
            long longestID = 0; // Track the highest ID found in the existing Excel sheets

            // Determine the latest ID from all sheets
            foreach (string sheet in sheetnamesList)
            {
                longestID = Math.Max(longestID, GetLatestId(ExcelFileName, sheet));
            }
            longestID++; // Increment to ensure new data has a unique ID

            // Define a list of closing price-related items
            List<string> closingList = ["priceMin", "priceMax", "priceYesterday", "priceFirst", "pClosing", "pDrCotVal", "zTotTran", "qTotTran5J", "qTotCap"];

            // Dictionary to store retrieved data from API responses
            Dictionary<int, Dictionary<string, string>> dataCollection = new Dictionary<int, Dictionary<string, string>>();

            await Parallel.ForAsync(0, urlParamsList.Count, async (i, cancellationToken) =>
            {
                var closingPriceData = new Dictionary<string, string> { { "SharedID", longestID.ToString() } };
                bool isDataValid = true; // Flag to track data validity

                // Fetch closing price data if required
                if (SelectedItems.Any(x => closingList.Contains(x)))
                {
                    var result = await GetClosingPriceInfo(urlParamsList[i], SelectedItems);
                    if (result != null && result.Count() > 0)
                    {
                        foreach (var item in result)
                        {
                            if (string.IsNullOrEmpty(item.Value))
                            {
                                isDataValid = false; // Mark as invalid if any value is empty
                            }
                            else if (NonZeroItems.Contains(item.Value))
                            {
                                // Ensure the value is non-zero if required
                                if (int.TryParse(item.Value, out int value) && value == 0)
                                    isDataValid = false;
                            }
                            else
                            {
                                closingPriceData.Add(item.Key, item.Value!);
                            }
                        }
                    }
                    else
                    {
                        isDataValid = false; // Mark as invalid if no data is received
                    }
                }

                // Fetch ETF data if required
                List<string> eftList = ["pRedTran", "pSubTran"];
                var etfData = new Dictionary<string, string>();
                if (SelectedItems.Any(x => eftList.Contains(x)) && isDataValid)
                    etfData = await GetETFByInsCode(urlParamsList[i]);

                // Merge ETF data into the existing data set
                if (etfData != null && etfData.Count > 0)
                {
                    foreach (var item in etfData)
                    {
                        if (string.IsNullOrEmpty(item.Value))
                        {
                            isDataValid = false; // Mark as invalid if any value is empty
                        }
                        else if (NonZeroItems.Contains(item.Value))
                        {
                            if (int.TryParse(item.Value, out int value) && value == 0)
                                isDataValid = false;
                        }
                        else
                        {
                            closingPriceData[item.Key] = item.Value!;
                        }
                    }
                }
                else
                {
                    isDataValid = false; // Mark as invalid if ETF data is missing
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
    }
}
