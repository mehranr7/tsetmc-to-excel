using System.Text.Json;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

class Program
{
    static async Task Main()
    {
        // Load configuration from appsettings.json
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        var settings = GetUserInput(config);

        Console.WriteLine("\nPress 'Q' to stop the data fetching process.\n");

        using var cts = new CancellationTokenSource();
        Task.Run(() =>
        {
            while (Console.ReadKey(true).Key != ConsoleKey.Q) { }
            cts.Cancel();
        });

        // Continuously fetch data until the user presses 'Q'
        while (!cts.Token.IsCancellationRequested)
        {
            await UpdateData(settings);
            await Task.Delay(settings.UpdateInterval * 1000, cts.Token);
        }

        Console.WriteLine("Data fetching stopped.");
    }

    /// <summary>
    /// Prompts the user for input, using default values from appsettings.json if skipped.
    /// </summary>
    static (string UrlParam, string FileName, string SheetName, int UpdateInterval, List<string> SelectedItems) GetUserInput(IConfiguration config)
    {
        Console.Write($"Enter the parameter for URL (Default: {config["ApiParameter"]}): ");
        string urlParam = Console.ReadLine().Trim();
        if (string.IsNullOrWhiteSpace(urlParam)) urlParam = config["ApiParameter"];

        Console.Write($"Enter the name of the Excel file (Default: {config["ExcelFileName"]}): ");
        string fileName = Console.ReadLine().Trim();
        if (string.IsNullOrWhiteSpace(fileName)) fileName = config["ExcelFileName"];
        if (!fileName.EndsWith(".xlsx")) fileName += ".xlsx";

        Console.Write($"Enter the name of the worksheet (Default: {config["SheetName"]}): ");
        string sheetName = Console.ReadLine().Trim();
        if (string.IsNullOrWhiteSpace(sheetName)) sheetName = config["SheetName"];

        Console.Write($"Enter the update interval in seconds (Default: {config["UpdateInterval"]}): ");
        string updateIntervalInput = Console.ReadLine().Trim();
        int updateInterval = string.IsNullOrWhiteSpace(updateIntervalInput) ? int.Parse(config["UpdateInterval"]) : int.Parse(updateIntervalInput);

        Console.WriteLine("Select the data fields (comma-separated or press Enter for default):");
        string[] options = { "priceMin", "priceMax", "priceYesterday", "priceFirst", "pClosing", "pDrCotVal", "zTotTran", "qTotTran5J", "qTotCap" };
        Console.WriteLine(string.Join(", ", options));

        List<string> allItems = config.GetSection("AllItems").Get<List<string>>();
        List<string> selectedItems = SelectItems(allItems);

        return (urlParam, fileName, sheetName, updateInterval, selectedItems);
    }

    /// <summary>
    /// Sends a request to the ClosingPrice API and returns a dictionary with selected values.
    /// </summary>
    static async Task<Dictionary<string, object>> GetClosingPriceInfo(string urlParam, List<string> selectedItems)
    {
        Dictionary<string, object> data = new Dictionary<string, object>();

        try
        {
            using var client = new HttpClient();
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
                    data[item] = value.ValueKind == JsonValueKind.Number ? value.GetDecimal() : value.ToString();
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
    static async Task<Dictionary<string, object>> GetETFByInsCode(string urlParam)
    {
        Dictionary<string, object> data = new Dictionary<string, object>();

        try
        {
            using var client = new HttpClient();
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
                data["pRedTran"] = pRedTran.GetDecimal();

            if (root.TryGetProperty("pSubTran", out JsonElement pSubTran))
                data["pSubTran"] = pSubTran.GetDecimal();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ETFByInsCode: {ex.Message}");
            ex = ex.InnerException;
            while (ex != null)
            {
                Console.WriteLine($"InnerException: {ex.Message}");
                ex = ex.InnerException;
            }
        }

        return data;
    }

    /// <summary>
    /// Combines data from both APIs and saves to Excel.
    /// </summary>
    static async Task UpdateData((string UrlParam, string FileName, string SheetName, int UpdateInterval, List<string> SelectedItems) settings)
    {
        List<string> closingList = ["priceMin", "priceMax", "priceYesterday", "priceFirst", "pClosing", "pDrCotVal", "zTotTran", "qTotTran5J", "qTotCap"];
        var closingPriceData = new Dictionary<string, object>();
        if (settings.SelectedItems.Any(x=> closingList.Contains(x)))
            closingPriceData = await GetClosingPriceInfo(settings.UrlParam, settings.SelectedItems);

        List<string> eftList = ["pRedTran", "pSubTran"];
        var etfData = new Dictionary<string, object>();
        if (settings.SelectedItems.Any(x => eftList.Contains(x)))
            etfData = await GetETFByInsCode(settings.UrlParam);

        // Merge dictionaries
        foreach (var kvp in etfData)
        {
            closingPriceData[kvp.Key] = kvp.Value;
        }

        if (closingPriceData.Any())
        {
            SaveToExcel(settings.FileName, settings.SheetName, closingPriceData);
        }
        else
        {
            Console.WriteLine("No data received to save!");
        }
    }

    /// <summary>
    /// Saves data to an Excel file, creating missing columns if necessary.
    /// </summary>
    static void SaveToExcel(string filePath, string sheetName, Dictionary<string, object> data)
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
            Console.WriteLine("Data updated and saved");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error savig excel: {ex.Message}");
            ex = ex.InnerException;
            while (ex != null)
            {
                Console.WriteLine($"InnerException: {ex.Message}");
                ex = ex.InnerException;
            }
        }
    }

    static List<string> SelectItems(List<string> items)
    {
        HashSet<int> selectedIndices = new HashSet<int>(Enumerable.Range(0, items.Count));
        ConsoleKey key;
        int currentIndex = 0;

        do
        {
            Console.Clear();
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
