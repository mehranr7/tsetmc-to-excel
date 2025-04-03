using Microsoft.Extensions.Configuration;
using static TseTmcToExcel.Excel.ExcelTools;
using static TseTmcToExcel.TseTmc.TseTmcTools;

namespace TseTmcToExcel
{
    public static class Program
    {
        public static async Task Main()
        {
            while (true)
            {
                // Load configuration settings from appsettings.json
                IConfiguration config = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // Retrieve user input or default settings from configuration
                var settings = GetUserInput(config);
                var counter = 0;

                Console.Clear();
                Console.WriteLine("Press 'Q' to stop the data fetching process.\n");

                // Create a cancellation token to handle user exit request
                using var cts = new CancellationTokenSource();
                _ = Task.Run(() =>
                {
                    while (Console.ReadKey(true).Key != ConsoleKey.Q) { }
                    cts.Cancel();
                });

                // Start continuous data fetching loop until 'Q' is pressed
                while (!cts.Token.IsCancellationRequested)
                {
                    counter++;
                    Console.WriteLine("______________________");
                    Console.WriteLine($"#{counter} Started at {DateTime.Now.ToString("HH:mm:ss")}\n");

                    _ = UpdateData(settings); // Fetch and update data asynchronously
                    await Task.Delay(settings.UpdateInterval * 1000, cts.Token); // Wait before fetching data again

                    Console.Clear();
                    Console.WriteLine("Press 'Q' to stop the data fetching process.\n");
                }

                Console.WriteLine("Data fetching stopped.\n");
            }
        }

        /// <summary>
        /// Retrieves user input for configuration settings, using defaults from appsettings.json if skipped.
        /// </summary>
        static (string UrlParam, string FileName, string SheetName, int UpdateInterval, List<string> NonZeroItems, List<string> SelectedItems) GetUserInput(IConfiguration config)
        {
            Console.Clear();
            Console.Write($"Enter the parameter for URL (Default: {config["ApiParameter"]}): ");
            string urlParam = Console.ReadLine()!.Trim();
            if (string.IsNullOrWhiteSpace(urlParam)) urlParam = config["ApiParameter"]!;

            Console.Clear();
            Console.Write($"Enter the name of the Excel file (Default: {config["ExcelFileName"]}): ");
            string fileName = Console.ReadLine()!.Trim();
            if (string.IsNullOrWhiteSpace(fileName)) fileName = config["ExcelFileName"]!;
            if (!fileName!.EndsWith(".xlsx")) fileName += ".xlsx";

            Console.Clear();
            Console.Write($"Enter the name of the worksheet (Default: {config["SheetName"]}): ");
            string sheetName = Console.ReadLine()!.Trim();
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = config["SheetName"]!;

            Console.Clear();
            Console.Write($"Enter the update interval in seconds (Default: {config["UpdateInterval"]}): ");
            string updateIntervalInput = Console.ReadLine()!.Trim();
            int updateInterval = string.IsNullOrWhiteSpace(updateIntervalInput) ? int.Parse(config["UpdateInterval"]!) : int.Parse(updateIntervalInput);

            Console.Clear();
            // Retrieve list of all available items from configuration
            List<string> allItems = config.GetSection("AllItems").Get<List<string>>()!;
            // Prompt user to select the items they want to fetch
            List<string> selectedItems = SelectItems(allItems, true, "Select the items needed to fetch.");

            Console.Clear();
            // Retrieve list of items that should not be zero
            List<string> nonZeroItems = config.GetSection("NonZeroItems").Get<List<string>>()!;
            // Prompt user to select non-zero constraint items
            List<string> selectedNonZeroItems = SelectItems(allItems, false, "Select the items that should not be zero.");

            return (urlParam, fileName, sheetName, updateInterval, selectedNonZeroItems, selectedItems);
        }

        /// <summary>
        /// Fetches data from APIs, validates it, and saves it to an Excel file.
        /// </summary>
        public static async Task UpdateData((string UrlParams, string FileName, string SheetNames, int UpdateInterval, List<string> NonZeroItems, List<string> SelectedItems) settings)
        {
            // Split multiple parameters into lists
            List<string> urlParamsList = settings.UrlParams.Replace(" ", "").Split(',').ToList();
            List<string> sheetnamesList = settings.SheetNames.Replace(" ", "").Split(',').ToList();

            // Semaphore to prevent concurrent issues
            SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);
            long longestID = 0;

            // Determine the latest ID from all sheets
            foreach (string sheet in sheetnamesList)
            {
                longestID = Math.Max(longestID, GetLatestId(settings.FileName, sheet));
            }
            longestID++;

            // Define list of closing price related items
            List<string> closingList = ["priceMin", "priceMax", "priceYesterday", "priceFirst", "pClosing", "pDrCotVal", "zTotTran", "qTotTran5J", "qTotCap"];

            // Dictionary to store retrieved data
            Dictionary<int, Dictionary<string, string>> dataCollection = new Dictionary<int, Dictionary<string, string>>();

            await Parallel.ForAsync(0, urlParamsList.Count, async (i, cancellationToken) =>
            {
                var closingPriceData = new Dictionary<string, string> { { "SharedID", longestID.ToString() } };
                bool isDataValid = true;

                // Fetch closing price data if required
                if (settings.SelectedItems.Any(x => closingList.Contains(x)))
                {
                    var result = await GetClosingPriceInfo(urlParamsList[i], settings.SelectedItems);
                    if (result != null && result.Count() > 0)
                    {
                        foreach (var item in result)
                        {
                            if (string.IsNullOrEmpty(item.Value))
                            {
                                isDataValid = false;
                            }
                            else if (settings.NonZeroItems.Contains(item.Value))
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
                        isDataValid = false;
                    }
                }

                // Fetch ETF data if required
                List<string> eftList = ["pRedTran", "pSubTran"];
                var etfData = new Dictionary<string, string>();
                if (settings.SelectedItems.Any(x => eftList.Contains(x)) && isDataValid)
                    etfData = await GetETFByInsCode(urlParamsList[i]);

                // Merge ETF data
                if (etfData != null && etfData.Count > 0)
                {
                    foreach (var item in etfData)
                    {
                        if (string.IsNullOrEmpty(item.Value))
                        {
                            isDataValid = false;
                        }
                        else if (settings.NonZeroItems.Contains(item.Value))
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
                    isDataValid = false;
                }

                // Store valid data
                if (isDataValid)
                    dataCollection.Add(i, closingPriceData);
            });

            // Save valid data to Excel
            if (dataCollection.Count == urlParamsList.Count)
            {
                foreach (var data in dataCollection)
                    SaveToExcel(settings.FileName, sheetnamesList[data.Key], data.Value);
            }
            else
            {
                Console.WriteLine("No valid data received to save!");
            }
        }

        /// <summary>
        /// Allows user to select items from a list using keyboard navigation.
        /// </summary>
        static List<string> SelectItems(List<string> items, bool isSelected, string prompt)
        {
            HashSet<int> selectedIndices = isSelected ? new HashSet<int>(Enumerable.Range(0, items.Count)) : new HashSet<int>();
            ConsoleKey key;
            int currentIndex = 0;

            do
            {
                Console.Clear();
                Console.WriteLine(prompt);
                Console.WriteLine("Use Arrow Keys to Navigate, Space to Toggle, Enter to Confirm:");

                for (int i = 0; i < items.Count; i++)
                {
                    string prefix = selectedIndices.Contains(i) ? "[X]" : "[ ]";
                    if (i == currentIndex)
                        Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"{prefix} {items[i]}");
                    Console.ResetColor();
                }

                key = Console.ReadKey(true).Key;

                switch (key)
                {
                    case ConsoleKey.UpArrow:
                        if (currentIndex > 0) currentIndex--;
                        break;
                    case ConsoleKey.DownArrow:
                        if (currentIndex < items.Count - 1) currentIndex++;
                        break;
                    case ConsoleKey.Spacebar:
                        if (selectedIndices.Contains(currentIndex))
                            selectedIndices.Remove(currentIndex);
                        else
                            selectedIndices.Add(currentIndex);
                        break;
                }
            } while (key != ConsoleKey.Enter);

            return selectedIndices.Select(i => items[i]).ToList();
        }
    }
}
