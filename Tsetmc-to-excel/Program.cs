using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

class Program
{
    static async Task Main()
    {
        while (true)
        {
            // Load configuration from appsettings.json
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var settings = GetUserInput(config);

            var counter = 0;

            Console.Clear();
            Console.WriteLine("Press 'Q' to stop the data fetching process.\n");

            using var cts = new CancellationTokenSource();
            _ = Task.Run(() =>
            {
                while (Console.ReadKey(true).Key != ConsoleKey.Q) { }
                cts.Cancel();
            });
            // Continuously fetch data until the user presses 'Q'
            while (!cts.Token.IsCancellationRequested)
            {
                counter++;
                Console.WriteLine("______________________");
                Console.WriteLine($"#{counter} Started at {DateTime.Now.ToString("HH:mm:ss")}\n");

                _ = UpdateData(settings);
                await Task.Delay(settings.UpdateInterval * 1000, cts.Token);
                Console.Clear();
                Console.WriteLine("Press 'Q' to stop the data fetching process.\n");
            }
            Console.WriteLine("Data fetching stopped.\n");
        }
    }

    /// <summary>
    /// Prompts the user for input, using default values from appsettings.json if skipped.
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
        List<string> allItems = config.GetSection("AllItems").Get<List<string>>()!;
        List<string> selectedItems = SelectItems(allItems, true, "Select the items are needed to fetch.");

        Console.Clear();
        List<string> nonZeroItems = config.GetSection("NonZeroItems").Get<List<string>>()!;
        List<string> selectedNonZeroItems = SelectItems(allItems, false, "Select the items are not allowed to be 0.");

        return (urlParam, fileName, sheetName, updateInterval, selectedNonZeroItems, selectedItems);
    }

    /// <summary>
    /// Sends a request to the ClosingPrice API and returns a dictionary with selected values.
    /// </summary>
    static async Task<Dictionary<string, string>> GetClosingPriceInfo(string urlParam, List<string> selectedItems)
    {
        Dictionary<string, string> data = new Dictionary<string, string>();

        // Load configuration from appsettings.json
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        try
        {

            using var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(int.Parse(config["UpdateInterval"]!));
            var request = new HttpRequestMessage(HttpMethod.Get, $"https://cdn.tsetmc.com/api/ClosingPrice/GetClosingPriceInfo/{urlParam}");

            request.Headers.Add("Accept", "application/json, text/plain, */*");
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            string jsonResponse = await response.Content.ReadAsStringAsync();

            using JsonDocument doc = JsonDocument.Parse(jsonResponse);
            var root = doc.RootElement.GetProperty("closingPriceInfo");

            // Add timestamp
            data["Date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Extract selected fields
            foreach (var item in selectedItems)
            {
                if (root.TryGetProperty(item, out JsonElement value))
                {
                    data[item] = value.ValueKind == JsonValueKind.Number ? value.ToString() : value.ToString();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ClosingPriceInfo: {ex.Message}");
        }

        return data;
    }

    /// <summary>
    /// Sends a request to the ETF API and returns a dictionary with pRedTran and pSubTran values.
    /// </summary>
    static async Task<Dictionary<string, string>> GetETFByInsCode(string urlParam)
    {
        Dictionary<string, string> data = new Dictionary<string, string>();

        // Load configuration from appsettings.json
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        try
        {
            using var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(int.Parse(config["UpdateInterval"]!));
            var request = new HttpRequestMessage(HttpMethod.Get, $"https://cdn.tsetmc.com/api/Fund/GetETFByInsCode/{urlParam}");

            request.Headers.Add("Accept", "application/json, text/plain, */*");
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            string jsonResponse = await response.Content.ReadAsStringAsync();

            using JsonDocument doc = JsonDocument.Parse(jsonResponse);
            var root = doc.RootElement.GetProperty("etf");

            // Extract pRedTran and pSubTran
            if (root.TryGetProperty("pRedTran", out JsonElement pRedTran))
                data["pRedTran"] = pRedTran.ToString();

            if (root.TryGetProperty("pSubTran", out JsonElement pSubTran))
                data["pSubTran"] = pSubTran.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ETFByInsCode: {ex.Message}");
            ex = ex.InnerException!;
            while (ex != null)
            {
                Console.WriteLine($"InnerException: {ex.Message}");
                ex = ex.InnerException!;
            }
        }

        return data;
    }

    /// <summary>
    /// Combines data from both APIs and saves to Excel.
    /// </summary>
    static async Task UpdateData((string UrlParams, string FileName, string SheetNames, int UpdateInterval,List<string> NonZeroItems, List<string> SelectedItems) settings)
    {
        List<string> urlParamsList = settings.UrlParams.Replace(" ","").Split(',').ToList();
        List<string> sheetnamesList = settings.SheetNames.Replace(" ", "").Split(',').ToList();
        SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);
        long longestID = 0;
        foreach (string sheet in sheetnamesList)
        {
            longestID = Math.Max(longestID, GetLatestId(settings.FileName, sheet));
        }
        longestID++;

        List<string> closingList = ["priceMin", "priceMax", "priceYesterday", "priceFirst", "pClosing", "pDrCotVal", "zTotTran", "qTotTran5J", "qTotCap"];

        Dictionary<int, Dictionary<string, string>> dataCollection = new Dictionary<int, Dictionary<string, string>>();
        await Parallel.ForAsync(0, urlParamsList.Count, async (i, cancellationToken) =>
        {
            var closingPriceData = new Dictionary<string, string>
            {
                { "SharedID", longestID.ToString() }
            };
            
            bool isDataValid = true;

            if (settings.SelectedItems.Any(x => closingList.Contains(x)))
            {
                var result = await GetClosingPriceInfo(urlParamsList[i], settings.SelectedItems);
                if(result != null && result.Count() > 0)
                {
                    foreach (var item in result)
                    {
                        if (item.Value == null || string.IsNullOrEmpty(item.Value))
                        {
                            isDataValid = false;
                        }
                        else if (settings.NonZeroItems.Contains(item.Value))
                        {
                            try
                            {
                                var value = int.Parse(item.Value);
                                if (value == 0)
                                    isDataValid = false;
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("The recieved data is not a number to check if its zero or not!");
                            }
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

            List<string> eftList = ["pRedTran", "pSubTran"];
            var etfData = new Dictionary<string, string>();
            if (settings.SelectedItems.Any(x => eftList.Contains(x)) && isDataValid)
                etfData = await GetETFByInsCode(urlParamsList[i]);

            // Merge dictionaries
            if (etfData != null && etfData.Count > 0)
            {
                foreach (var item in etfData)
                {
                    if (item.Value == null || string.IsNullOrEmpty(item.Value))
                    {
                        isDataValid = false;
                    }
                    else if (settings.NonZeroItems.Contains(item.Value))
                    {
                        try
                        {
                            var value = int.Parse(item.Value);
                            if (value == 0)
                                isDataValid = false;
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("The recieved data is not a number to check if its zero or not!");
                        }
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

            if(isDataValid)
                dataCollection.Add(i, closingPriceData);
        });


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
    /// Saves data to an Excel file, creating missing columns if necessary.
    /// </summary>
    static void SaveToExcel(string filePath, string sheetName, Dictionary<string, string> data)
    {
        try
        {
            FileInfo file = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = file.Exists ? new ExcelPackage(file) : new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName) ?? package.Workbook.Worksheets.Add(sheetName);

            // Ensure all columns exist
            List<string> existingHeaders = new List<string>();
            if (worksheet.Dimension != null)
            {
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    existingHeaders.Add(worksheet.Cells[1, col].Text);
                }
            }

            int colIndex = 1;
            foreach (var key in data.Keys)
            {
                if (!existingHeaders.Contains(key))
                {
                    worksheet.Cells[1, existingHeaders.Count + 1].Value = key;
                    existingHeaders.Add(key);
                }
                colIndex++;
            }

            int newRow = worksheet.Dimension?.Rows + 1 ?? 2;
            colIndex = 1;
            foreach (var key in existingHeaders)
            {
                if (data.ContainsKey(key))
                {
                    worksheet.Cells[newRow, colIndex].Value = data[key];
                }
                colIndex++;
            }

            package.SaveAs(file);
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}\t{sheetName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error savig excel: {ex.Message}");
            ex = ex.InnerException!;
            while (ex != null)
            {
                Console.WriteLine($"InnerException: {ex.Message}");
                ex = ex.InnerException!;
            }
        }
    }

    static long GetLatestId(string filePath, string sheetName, int columnIndex = 1)
    {
        try
        {
            FileInfo file = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(file);
            var worksheet = package.Workbook.Worksheets[sheetName];

            // No data in sheet
            if (worksheet == null || worksheet.Dimension == null)
                return 0;

            int lastRow = worksheet.Dimension.Rows;
            long latestId = 0;

            // Assuming row 1 is header
            for (int row = 2; row <= lastRow; row++)
            {
                var cellValue = worksheet.Cells[row, columnIndex].Value;
                if (cellValue != null && long.TryParse(cellValue.ToString(), out long id))
                {
                    latestId = Math.Max(latestId, id);
                }
            }

            return latestId;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving latest ID: {ex.Message}");
            return 0;
        }
    }

    static List<string> SelectItems(List<string> items, bool isSelected , string prompt)
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
