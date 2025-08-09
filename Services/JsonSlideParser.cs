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
                    LayoutType = ParseLayoutType(jsonSlide.Layout)
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
        /// Parses JSON slide content from a string with embedded base64 images
        /// </summary>
        /// <param name="jsonContent">JSON content as string</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <returns>Parsed presentation content with images from base64 data</returns>
        public static PresentationContent ParseFromStringWithEmbeddedImages(
            string jsonContent,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant")
        {
            // Parse JSON
            var jsonSlideContent = JsonSerializer.Deserialize<JsonSlideContent>(jsonContent, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (jsonSlideContent?.Slides == null || !jsonSlideContent.Slides.Any())
                throw new InvalidOperationException("No slides found in JSON content");

            // Create a temporary directory for images
            var tempDir = Path.Combine(Path.GetTempPath(), "PowerPointGenerator", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            // Create dictionary for image lookup
            var imageDict = new Dictionary<string, string>();
            if (jsonSlideContent.Images != null)
            {
                foreach (var jsonImage in jsonSlideContent.Images)
                {
                    if (!string.IsNullOrWhiteSpace(jsonImage.Data))
                    {
                        try
                        {
                            // Extract base64 data (remove data:image/type;base64, prefix if present)
                            var base64Data = jsonImage.Data;
                            string extension = ".png"; // Default extension
                            
                            if (base64Data.StartsWith("data:image/"))
                            {
                                var commaIndex = base64Data.IndexOf(',');
                                if (commaIndex >= 0)
                                {
                                    // Extract MIME type to determine extension
                                    var mimeTypePart = base64Data.Substring(0, commaIndex);
                                    extension = DetermineExtensionFromMimeType(mimeTypePart);
                                    base64Data = base64Data.Substring(commaIndex + 1);
                                }
                            }
                            else
                            {
                                // Check if ID already has extension
                                var idExtension = Path.GetExtension(jsonImage.Id);
                                if (!string.IsNullOrWhiteSpace(idExtension))
                                    extension = idExtension;
                            }

                            // Convert base64 to bytes
                            var imageBytes = Convert.FromBase64String(base64Data);
                            
                            // Create filename with proper extension
                            var fileName = Path.GetFileNameWithoutExtension(jsonImage.Id) + extension;
                            var imagePath = Path.Combine(tempDir, fileName);
                            
                            // Write image to temporary file
                            File.WriteAllBytes(imagePath, imageBytes);
                            imageDict[jsonImage.Id] = imagePath;
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException($"Failed to process image {jsonImage.Id}: {ex.Message}", ex);
                        }
                    }
                }
            }

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
                    LayoutType = ParseLayoutType(jsonSlide.Layout)
                };

                // Parse image from suggested_image field
                if (!string.IsNullOrWhiteSpace(jsonSlide.SuggestedImage))
                {
                    var imageId = ExtractImageIdFromSuggestion(jsonSlide.SuggestedImage);
                    if (!string.IsNullOrWhiteSpace(imageId) && imageDict.ContainsKey(imageId))
                    {
                        slideContent.Images.Add(new ImageContent
                        {
                            FilePath = imageDict[imageId],
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
        /// Extracts the image ID from the suggested_image field
        /// </summary>
        /// <param name="suggestedImage">The suggested image text (e.g., "Use Image 1: original-1753723327847.png")</param>
        /// <returns>Image ID</returns>
        private static string ExtractImageIdFromSuggestion(string suggestedImage)
        {
            // Extract image ID from patterns like "Use Image 1: original-1753723327847.png"
            var patterns = new[]
            {
                @"Use Image \d+:\s*""?([^""]+)""?",  // "Use Image 1: "id""
                @"Use Image \d+:\s*([^""]+)",        // Use Image 1: id
                @"""([^""]+)""",                     // "id"
                @"([^""]+)"                          // id
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(suggestedImage, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Determines the image file extension based on MIME type string
        /// </summary>
        /// <param name="mimeTypePart">MIME type part from data URL (e.g., "data:image/jpeg;base64")</param>
        /// <returns>File extension including the dot</returns>
        private static string DetermineExtensionFromMimeType(string mimeTypePart)
        {
            if (string.IsNullOrWhiteSpace(mimeTypePart))
                return ".png"; // Default

            // Extract the image type from MIME type
            if (mimeTypePart.Contains("image/"))
            {
                var imageType = mimeTypePart.Substring(mimeTypePart.IndexOf("image/") + 6);
                if (imageType.Contains(";"))
                    imageType = imageType.Substring(0, imageType.IndexOf(";"));

                return imageType.ToLowerInvariant() switch
                {
                    "jpeg" or "jpg" => ".jpg",
                    "png" => ".png",
                    "gif" => ".gif",
                    "bmp" => ".bmp",
                    "webp" => ".webp",
                    "svg+xml" => ".svg",
                    _ => ".png" // Default
                };
            }

            return ".png"; // Default extension
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

        /// <summary>
        /// Parses layout type from string
        /// </summary>
        /// <param name="layoutString">Layout string from JSON</param>
        /// <returns>Corresponding SlideLayoutType</returns>
        private static SlideLayoutType ParseLayoutType(string layoutString)
        {
            if (string.IsNullOrWhiteSpace(layoutString))
                return SlideLayoutType.ImageFocused; // Default

            return layoutString.ToLowerInvariant() switch
            {
                "title" => SlideLayoutType.Title,
                "titleandcontent" or "title_and_content" or "title-and-content" => SlideLayoutType.TitleAndContent,
                "imagefocused" or "image_focused" or "image-focused" => SlideLayoutType.ImageFocused,
                "imagegrid" or "image_grid" or "image-grid" => SlideLayoutType.ImageGrid,
                "singleimagewithcaption" or "single_image_with_caption" or "single-image-with-caption" => SlideLayoutType.SingleImageWithCaption,
                "twoimagecomparison" or "two_image_comparison" or "two-image-comparison" => SlideLayoutType.TwoImageComparison,
                "productshowcase" or "product_showcase" or "product-showcase" => SlideLayoutType.ProductShowcase,
                _ => SlideLayoutType.ImageFocused // Default fallback
            };
        }
    }
}
