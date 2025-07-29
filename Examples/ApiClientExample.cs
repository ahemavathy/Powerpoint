using PowerPointGenerator.Client;
using System.Text.Json;

namespace PowerPointGenerator.Examples
{
    /// <summary>
    /// Example of how to use the PowerPoint API client
    /// </summary>
    public class ApiClientExample
    {
        public static async Task RunExampleAsync()
        {
            // Create API client
            var client = new PowerPointApiClient("http://localhost:5000");

            try
            {
                // Check if API is healthy
                Console.WriteLine("Checking API health...");
                var isHealthy = await client.IsHealthyAsync();
                if (!isHealthy)
                {
                    Console.WriteLine("❌ API is not available. Make sure the web service is running.");
                    return;
                }
                Console.WriteLine("✅ API is healthy");

                // Create presentation from JSON
                var jsonContent = @"{
  ""slides"": [
    {
      ""title"": ""Welcome to Our API Demo"",
      ""description"": ""This presentation was created using our PowerPoint Generator Web API."",
      ""suggested_image"": ""demo_image_1.png""
    },
    {
      ""title"": ""Key Features"",
      ""description"": ""• Generate presentations from JSON\n• RESTful API endpoints\n• Download generated files\n• Cross-platform compatible"",
      ""suggested_image"": ""demo_image_2.png""
    },
    {
      ""title"": ""Easy Integration"",
      ""description"": ""Call our API from any programming language or platform that supports HTTP requests."",
      ""suggested_image"": ""demo_image_3.png""
    }
  ]
}";

                Console.WriteLine("\n🔄 Creating presentation...");
                var response = await client.CreatePresentationAsync(
                    jsonContent,
                    presentationName: "API_Demo_Presentation",
                    presentationTitle: "PowerPoint API Demo",
                    author: "API Client Example"
                );

                if (response != null && response.Success)
                {
                    Console.WriteLine("✅ Presentation created successfully!");
                    Console.WriteLine($"   File: {response.FileName}");
                    Console.WriteLine($"   Size: {response.FileSize:N0} bytes");
                    Console.WriteLine($"   Slides: {response.SlideCount}");
                    Console.WriteLine($"   Download URL: {response.DownloadUrl}");

                    // Download the file
                    Console.WriteLine("\n📥 Downloading presentation...");
                    var fileData = await client.DownloadPresentationAsync(response.FileName);
                    var localPath = Path.Combine(Environment.CurrentDirectory, "Downloaded_" + response.FileName);
                    await File.WriteAllBytesAsync(localPath, fileData);
                    Console.WriteLine($"✅ Downloaded to: {localPath}");
                }

                // List all presentations
                Console.WriteLine("\n📋 Getting list of presentations...");
                var presentations = await client.GetPresentationListAsync();
                if (presentations != null && presentations.Any())
                {
                    Console.WriteLine($"Found {presentations.Count} presentation(s):");
                    foreach (var pres in presentations.Take(5)) // Show first 5
                    {
                        Console.WriteLine($"  • {pres.FileName} ({pres.FileSize:N0} bytes) - {pres.CreatedAt:yyyy-MM-dd HH:mm}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
            }
            finally
            {
                client.Dispose();
            }
        }
    }
}
