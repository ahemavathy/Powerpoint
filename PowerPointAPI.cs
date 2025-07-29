using PowerPointGenerator.Models;
using PowerPointGenerator.Services;

namespace PowerPointGenerator
{
    /// <summary>
    /// Simple API for creating PowerPoint presentations from structured content
    /// </summary>
    public static class PowerPointAPI
    {
        /// <summary>
        /// Creates a PowerPoint presentation from structured slide content text
        /// </summary>
        /// <param name="slideContentText">Structured slide content in the format with ### Slide X:</param>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path where images are located (defaults to ./Images)</param>
        /// <returns>Path to the created presentation file</returns>
        public static async Task<string> CreatePresentationFromText(
            string slideContentText,
            string outputPath,
            string presentationTitle = "AI Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            // Use default image path if not provided
            imageBasePath ??= Path.Combine(Environment.CurrentDirectory, "Images");

            // Parse the structured content
            var presentationContent = SlideContentParser.ParseSlideContent(
                slideContentText,
                presentationTitle,
                author,
                imageBasePath
            );

            // Generate the presentation
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(presentationContent, outputPath);

            return outputPath;
        }

        /// <summary>
        /// Creates a PowerPoint presentation from PresentationContent object
        /// </summary>
        /// <param name="content">Presentation content object</param>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        /// <returns>Path to the created presentation file</returns>
        public static async Task<string> CreatePresentation(PresentationContent content, string outputPath)
        {
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(content, outputPath);
            return outputPath;
        }

        /// <summary>
        /// Creates a PowerPoint presentation from JSON slide content file
        /// </summary>
        /// <param name="jsonFilePath">Path to the JSON file containing slide content</param>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path where images are located (defaults to ./Images)</param>
        /// <returns>Path to the created presentation file</returns>
        public static async Task<string> CreatePresentationFromJsonFile(
            string jsonFilePath,
            string outputPath,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            // Use default image path if not provided
            imageBasePath ??= Path.Combine(Environment.CurrentDirectory, "Images");

            // Parse the JSON content
            var presentationContent = JsonSlideParser.ParseFromFile(
                jsonFilePath, presentationTitle, author, imageBasePath);

            // Create the presentation
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(presentationContent, outputPath);

            return outputPath;
        }

    }
}
