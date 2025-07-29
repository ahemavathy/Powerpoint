using System.Text;
using System.Text.Json;

namespace PowerPointGenerator.Client
{
    /// <summary>
    /// Example client for calling the PowerPoint Generator Web API
    /// </summary>
    public class PowerPointApiClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;

        public PowerPointApiClient(string baseUrl = "http://localhost:5000")
        {
            _httpClient = new HttpClient();
            _baseUrl = baseUrl.TrimEnd('/');
        }

        /// <summary>
        /// Creates a presentation from JSON content
        /// </summary>
        public async Task<PresentationApiResponse?> CreatePresentationAsync(string jsonContent, 
            string? presentationName = null, string? presentationTitle = null, string? author = null)
        {
            var request = new
            {
                JsonContent = jsonContent,
                PresentationName = presentationName,
                PresentationTitle = presentationTitle,
                Author = author
            };

            var json = JsonSerializer.Serialize(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync($"{_baseUrl}/api/presentation/create-from-json", content);
            
            if (response.IsSuccessStatusCode)
            {
                var responseJson = await response.Content.ReadAsStringAsync();
                return JsonSerializer.Deserialize<PresentationApiResponse>(responseJson, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
            }

            var errorContent = await response.Content.ReadAsStringAsync();
            throw new Exception($"API Error: {response.StatusCode} - {errorContent}");
        }

        /// <summary>
        /// Downloads a presentation file
        /// </summary>
        public async Task<byte[]> DownloadPresentationAsync(string fileName)
        {
            var response = await _httpClient.GetAsync($"{_baseUrl}/api/presentation/download/{fileName}");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsByteArrayAsync();
        }

        /// <summary>
        /// Gets the list of available presentations
        /// </summary>
        public async Task<List<PresentationFileInfo>?> GetPresentationListAsync()
        {
            var response = await _httpClient.GetAsync($"{_baseUrl}/api/presentation/list");
            response.EnsureSuccessStatusCode();
            
            var json = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<List<PresentationFileInfo>>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }

        /// <summary>
        /// Checks if the API service is healthy
        /// </summary>
        public async Task<bool> IsHealthyAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/api/presentation/health");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    public class PresentationApiResponse
    {
        public bool Success { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string PresentationName { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public long FileSize { get; set; }
        public int SlideCount { get; set; }
        public string DownloadUrl { get; set; } = string.Empty;
    }

    public class PresentationFileInfo
    {
        public string FileName { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public long FileSize { get; set; }
        public string DownloadUrl { get; set; } = string.Empty;
    }
}
