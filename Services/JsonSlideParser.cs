using System.Text.Json;
using System.Text.RegularExpressions;
using PowerPointGenerator.Models;

namespace PowerPointGenerator.Services
{
    /// <summary>
    /// Parser for JSON slide content format
    /// </summary>
    public static class JsonSlideParser
    {
        /// <summary>
        /// Parses JSON slide content from a file
        /// </summary>
        /// <param name="jsonFilePath">Path to the JSON file</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Parsed presentation content</returns>
        public static PresentationContent ParseFromFile(
            string jsonFilePath,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"JSON file not found: {jsonFilePath}");

            var jsonContent = File.ReadAllText(jsonFilePath);
            return ParseFromString(jsonContent, presentationTitle, author, imageBasePath);
        }

        /// <summary>
        /// Parses JSON slide content from a string
        /// </summary>
        /// <param name="jsonContent">JSON content as string</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Parsed presentation content</returns>
        public static PresentationContent ParseFromString(
            string jsonContent,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            imageBasePath ??= Path.Combine(Environment.CurrentDirectory, "Images");

            // Parse JSON
            var jsonSlideContent = JsonSerializer.Deserialize<JsonSlideContent>(jsonContent, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (jsonSlideContent?.Slides == null || !jsonSlideContent.Slides.Any())
                throw new InvalidOperationException("No slides found in JSON content");

            // Convert to presentation content
            var presentationContent = new PresentationContent
            {
                Title = presentationTitle,
                Author = author
            };

            foreach (var jsonSlide in jsonSlideContent.Slides)
            {
                var slideContent = new SlideContent
                {
                    Title = jsonSlide.Title,
                    Description = jsonSlide.Description,
                    LayoutType = SlideLayoutType.ImageFocused
                };

                // Parse image from suggested_image field
                if (!string.IsNullOrWhiteSpace(jsonSlide.SuggestedImage))
                {
                    var imagePath = ExtractImagePathFromSuggestion(jsonSlide.SuggestedImage, imageBasePath);
                    if (!string.IsNullOrWhiteSpace(imagePath))
                    {
                        slideContent.Images.Add(new ImageContent
                        {
                            FilePath = imagePath,
                            AltText = jsonSlide.Title,
                            Caption = jsonSlide.Description
                        });
                    }
                }

                presentationContent.Slides.Add(slideContent);
            }

            return presentationContent;
        }

        /// <summary>
        /// Extracts the image filename from the suggested_image field
        /// </summary>
        /// <param name="suggestedImage">The suggested image text (e.g., "Use Image 1: filename.png")</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Full path to the image file</returns>
        private static string ExtractImagePathFromSuggestion(string suggestedImage, string imageBasePath)
        {
            // Extract filename from patterns like "Use Image 1: filename.png" or just "filename.png"
            var patterns = new[]
            {
                @"Use Image \d+:\s*""?([^""]+\.(?:png|jpg|jpeg|gif|bmp))""?",  // "Use Image 1: "filename.png""
                @"Use Image \d+:\s*([^""]+\.(?:png|jpg|jpeg|gif|bmp))",        // Use Image 1: filename.png
                @"""([^""]+\.(?:png|jpg|jpeg|gif|bmp))""",                     // "filename.png"
                @"([^""]+\.(?:png|jpg|jpeg|gif|bmp))"                          // filename.png
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(suggestedImage, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var filename = match.Groups[1].Value.Trim();
                    var fullPath = Path.Combine(imageBasePath, filename);
                    return fullPath;
                }
            }

            return string.Empty;
        }
    }
}
