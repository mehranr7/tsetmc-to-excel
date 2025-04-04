using System.Text.Json;

namespace TseTmcToExcel
{
    public static class TseTmcTools
    {
        /// <summary>
        /// Sends a request to the GetMarketOverview API and returns a dictionary with selected values.
        /// </summary>
        /// <param name="selectedItems">A list of JSON properties to extract from the response.</param>
        /// <returns>A dictionary containing selected key-value pairs from the API response.</returns>
        public static async Task<Dictionary<string, string>> GetMarketOverview()
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            try
            {
                var client = new HttpClient();
                // Set timeout based on the configured update interval
                client.Timeout = TimeSpan.FromSeconds(IO.Timeout);


                // Create the HTTP request to fetch Market Overview
                var request = new HttpRequestMessage(HttpMethod.Get, "https://cdn.tsetmc.com/api/MarketData/GetMarketOverview/1");

                // Send the request and ensure a successful response
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Read the JSON response
                var jsonResponse = await response.Content.ReadAsStringAsync();

                using JsonDocument doc = JsonDocument.Parse(jsonResponse);
                var root = doc.RootElement.GetProperty("marketOverview");

                // Add the current timestamp to the data
                data["Date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Extract the selected fields from the JSON response
                foreach (var item in IO.SelectedItems)
                {
                    if (root.TryGetProperty(item, out JsonElement value))
                    {
                        // Convert numbers to string representation
                        data[item] = value.ValueKind == JsonValueKind.Number ? value.ToString() : value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle errors by logging the message
                Console.WriteLine($"Error fetching GetMarketOverview: {ex.Message}");
            }
            return data;
        }

        /// <summary>
        /// Sends a request to the ClosingPrice API and returns a dictionary with selected values.
        /// </summary>
        /// <param name="urlParam">The URL parameter for the API request.</param>
        /// <param name="selectedItems">A list of JSON properties to extract from the response.</param>
        /// <returns>A dictionary containing selected key-value pairs from the API response.</returns>
        public static async Task<Dictionary<string, string>> GetClosingPriceInfo(string urlParam, List<string> selectedItems)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            try
            {
                using var client = new HttpClient();

                // Set timeout based on the configured update interval
                client.Timeout = TimeSpan.FromSeconds(IO.Timeout);

                // Create the HTTP request to fetch ClosingPriceInfo
                var request = new HttpRequestMessage(HttpMethod.Get, $"https://cdn.tsetmc.com/api/ClosingPrice/GetClosingPriceInfo/{urlParam}");

                // Add required headers to the request
                request.Headers.Add("Accept", "application/json, text/plain, */*");
                request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

                // Send the request and ensure a successful response
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Read the JSON response
                string jsonResponse = await response.Content.ReadAsStringAsync();

                using JsonDocument doc = JsonDocument.Parse(jsonResponse);
                var root = doc.RootElement.GetProperty("closingPriceInfo");

                // Add the current timestamp to the data
                data["Date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Extract the selected fields from the JSON response
                foreach (var item in selectedItems)
                {
                    if (root.TryGetProperty(item, out JsonElement value))
                    {
                        // Convert numbers to string representation
                        data[item] = value.ValueKind == JsonValueKind.Number ? value.ToString() : value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle errors by logging the message
                Console.WriteLine($"Error fetching ClosingPriceInfo: {ex.Message}");
            }

            return data;
        }

        /// <summary>
        /// Sends a request to the ETF API and returns a dictionary with pRedTran and pSubTran values.
        /// </summary>
        /// <param name="urlParam">The URL parameter for the API request.</param>
        /// <returns>A dictionary containing pRedTran and pSubTran values from the API response.</returns>
        public static async Task<Dictionary<string, string>> GetETFByInsCode(string urlParam)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            try
            {
                using var client = new HttpClient();

                // Set timeout based on the configured update interval
                client.Timeout = TimeSpan.FromSeconds(IO.Timeout);

                // Create the HTTP request to fetch ETF data
                var request = new HttpRequestMessage(HttpMethod.Get, $"https://cdn.tsetmc.com/api/Fund/GetETFByInsCode/{urlParam}");

                // Add required headers to the request
                request.Headers.Add("Accept", "application/json, text/plain, */*");
                request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

                // Send the request and ensure a successful response
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Read the JSON response
                string jsonResponse = await response.Content.ReadAsStringAsync();

                using JsonDocument doc = JsonDocument.Parse(jsonResponse);
                var root = doc.RootElement.GetProperty("etf");

                // Extract "pRedTran" and "pSubTran" values from the JSON response
                if (root.TryGetProperty("pRedTran", out JsonElement pRedTran))
                    data["pRedTran"] = pRedTran.ToString();

                if (root.TryGetProperty("pSubTran", out JsonElement pSubTran))
                    data["pSubTran"] = pSubTran.ToString();
            }
            catch (Exception ex)
            {
                // Handle errors by logging the message and any inner exceptions
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
    }
}
