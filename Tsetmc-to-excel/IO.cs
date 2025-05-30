﻿using Microsoft.Extensions.Configuration;

namespace TseTmcToExcel
{
    public static class IO
    {
        // Static properties to hold configuration values
        public static string ExcelFileName { get; private set; }
        public static string StockSheetName { get; private set; }
        public static string OverviewSheetName { get; private set; }
        public static int UpdateInterval { get; private set; }
        public static int Timeout { get; private set; }
        public static bool AskSettings { get; private set; }
        public static bool isParallel { get; private set; }
        public static List<string> ApiParameterList { get; private set; }
        public static List<string> StockNameList { get; private set; }
        public static List<string> ClosingItems { get; private set; }
        public static List<string> EtfItems { get; private set; }
        public static List<string> MarketOverviewItems { get; private set; }
        public static List<string> AllItems { get; private set; }
        public static List<string> SelectedItems { get; private set; }
        public static List<string> NonZeroItems { get; private set; }

        /// <summary>
        /// Initializes the configuration settings and prompts the user for inputs if required.
        /// </summary>
        public static void Initialize()
        {
            // Load configuration settings from appsettings.json
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory()) // Set base path to the current directory
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Load settings from JSON file
                .Build(); // Build the configuration

            // Retrieve values from configuration and assign them to static properties
            ExcelFileName = configuration.GetValue<string>("ExcelFileName")!;
            StockSheetName = configuration.GetValue<string>("StockSheetName")!;
            OverviewSheetName = configuration.GetValue<string>("OverviewSheetName")!;
            UpdateInterval = configuration.GetValue<int>("UpdateInterval");
            Timeout = configuration.GetValue<int>("Timeout");
            try
            {
                AskSettings = configuration.GetValue<bool>("AskSettings");
            }
            catch (Exception)
            {
                AskSettings = configuration.GetValue<int>("AskSettings") == 1;
            }
            try
            {
                isParallel = configuration.GetValue<bool>("Parallel");
            }
            catch (Exception)
            {
                isParallel = configuration.GetValue<int>("Parallel") == 1;
            }

            // Load lists from configuration file, ensuring they are not null
            ApiParameterList = configuration.GetSection("ApiParameter").Get<List<string>>() ?? new List<string>();
            StockNameList = configuration.GetSection("StockNameList").Get<List<string>>() ?? new List<string>();
            SelectedItems = configuration.GetSection("SelectedItems").Get<List<string>>() ?? new List<string>();
            NonZeroItems = configuration.GetSection("NonZeroItems").Get<List<string>>() ?? new List<string>();
            ClosingItems = configuration.GetSection("ClosingItems").Get<List<string>>() ?? new List<string>();
            EtfItems = configuration.GetSection("EtfItems").Get<List<string>>() ?? new List<string>();
            MarketOverviewItems = configuration.GetSection("MarketOverviewItems").Get<List<string>>() ?? new List<string>();

            // If AskSettings is enabled, prompt the user for custom input
            if (AskSettings)
            {
                Console.Clear();
                Console.Write($"Enter the list of parameters for the URL separated by a comma (Default: {string.Join(", ", ApiParameterList)}): ");
                string? apiParameterInput = Console.ReadLine()?.Trim();
                var apiToUse = string.IsNullOrWhiteSpace(apiParameterInput)
                    ? ApiParameterList
                    : apiParameterInput.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
                ApiParameterList = apiToUse;


                Console.Clear();
                Console.Write($"Enter the list of Stock Names in order separated by a comma (Default: {string.Join(", ", StockNameList)}): ");
                string? stockNameListInput = Console.ReadLine()?.Trim();
                var stockListToUse = string.IsNullOrWhiteSpace(stockNameListInput)
                    ? StockNameList
                    : stockNameListInput.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
                StockNameList = stockListToUse;

                Console.Clear();
                Console.Write($"Enter the name of the Excel file (Default: {ExcelFileName}): ");
                string fileName = Console.ReadLine()!.Trim();
                if (string.IsNullOrWhiteSpace(fileName)) fileName = ExcelFileName;
                ExcelFileName = fileName;

                Console.Clear();
                Console.Write($"Enter the name of the Stocks worksheet (Default: {StockSheetName}): ");
                string stockSheetName = Console.ReadLine()!.Trim();
                if (string.IsNullOrWhiteSpace(stockSheetName)) stockSheetName = StockSheetName;
                StockSheetName = stockSheetName;

                Console.Clear();
                Console.Write($"Enter the name of the Market Overview  worksheet (Default: {OverviewSheetName}): ");
                string overviewSheetName = Console.ReadLine()!.Trim();
                if (string.IsNullOrWhiteSpace(overviewSheetName)) overviewSheetName = OverviewSheetName;
                OverviewSheetName = overviewSheetName;

                Console.Clear();
                Console.Write($"Enter the update interval in seconds (Default: {UpdateInterval}): ");
                string updateIntervalInput = Console.ReadLine()!.Trim();
                int updateInterval = string.IsNullOrWhiteSpace(updateIntervalInput) ? UpdateInterval : int.Parse(updateIntervalInput);
                UpdateInterval = updateInterval;

                Console.Clear();
                Console.Write($"Enter the timeout of requests in seconds (Default: {Timeout}): ");
                string timeOutInput = Console.ReadLine()!.Trim();
                int timeout = string.IsNullOrWhiteSpace(timeOutInput) ? Timeout : int.Parse(timeOutInput);
                Timeout = timeout;

                var trueFalseUserInput = new List<string>(["y", "n", "0", "1"]);
                var parallelInput = "";
                while (!trueFalseUserInput.Any(x=>x.Equals(parallelInput)))
                {
                    Console.Clear();
                    Console.Write($"Do you want to send requsts using parallel? Y/N or 1/0 (Default: {isParallel}): ");
                    parallelInput = Console.ReadKey().KeyChar.ToString().ToLower();
                }
                isParallel = parallelInput!.Equals("y") || parallelInput.Equals("1");


                // Combine all of the available items
                AllItems = new List<string>();
                AllItems.AddRange(EtfItems);
                AllItems.AddRange(ClosingItems);
                AllItems.AddRange(MarketOverviewItems);

                Console.Clear();
                // Prompt the user to select items to fetch from the API
                SelectedItems = SelectItems(AllItems, SelectedItems, "Select the items needed to fetch.");

                Console.Clear();
                // Prompt the user to select items that should not have a zero value
                NonZeroItems = SelectItems(AllItems, NonZeroItems, "Select the items that should not be zero.");
            }

            // Ensure that the Excel file has the correct extension
            if (!ExcelFileName!.EndsWith(".xlsx")) ExcelFileName += ".xlsx";
        }

        /// <summary>
        /// Displays a selection menu for the user to choose items from a given list.
        /// Allows navigation using arrow keys and selection using the space bar.
        /// </summary>
        /// <param name="items">The list of available items.</param>
        /// <param name="selectedItems">The list of items pre-selected by the user.</param>
        /// <param name="prompt">The message displayed to the user before selection.</param>
        /// <returns>A list of selected items.</returns>
        private static List<string> SelectItems(List<string> items, List<string> selectedItems, string prompt)
        {
            // Store the indices of selected items
            HashSet<int> selectedIndices = new HashSet<int>();
            for (int i = 0; i < items.Count; i++)
            {
                if (selectedItems.Contains(items[i]))
                {
                    selectedIndices.Add(i); // Mark pre-selected items
                }
            }

            ConsoleKey key;
            int currentIndex = 0; // Track the currently highlighted item

            do
            {
                Console.Clear();
                Console.WriteLine(prompt);
                Console.WriteLine("Use Arrow Keys to Navigate, Space to Toggle, Enter to Confirm:");

                // Display the items with selection indicators
                for (int i = 0; i < items.Count; i++)
                {
                    string prefix = selectedIndices.Contains(i) ? "[X]" : "[ ]"; // Indicate selection status
                    if (i == currentIndex)
                        Console.ForegroundColor = ConsoleColor.Cyan; // Highlight the current selection
                    Console.WriteLine($"{prefix} {items[i]}");
                    Console.ResetColor();
                }

                key = Console.ReadKey(true).Key;

                // Handle navigation and selection input
                switch (key)
                {
                    case ConsoleKey.UpArrow:
                        if (currentIndex > 0) currentIndex--; // Move up
                        break;
                    case ConsoleKey.DownArrow:
                        if (currentIndex < items.Count - 1) currentIndex++; // Move down
                        break;
                    case ConsoleKey.Spacebar:
                        if (selectedIndices.Contains(currentIndex))
                            selectedIndices.Remove(currentIndex); // Deselect item
                        else
                            selectedIndices.Add(currentIndex); // Select item
                        break;
                }
            } while (key != ConsoleKey.Enter); // Exit loop when Enter key is pressed

            // Return the list of selected items based on user input
            return selectedIndices.Select(i => items[i]).ToList();
        }
    }
}