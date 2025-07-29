using PowerPointGenerator;

namespace PowerPointGenerator.Examples
{
    /// <summary>
    /// Example usage of JSON-based PowerPoint generation
    /// </summary>
    public static class JsonExample
    {
        /// <summary>
        /// Demonstrates creating a presentation from a JSON file
        /// </summary>
        public static async Task<string> CreateFromJsonFile()
        {
            var jsonFilePath = Path.Combine(Environment.CurrentDirectory, "slides_content.json");
            var outputPath = Path.Combine(Environment.CurrentDirectory, "JSON_Example_Presentation.pptx");

            return await PowerPointAPI.CreatePresentationFromJsonFile(
                jsonFilePath,
                outputPath,
                "Premium Can Opener - Product Showcase",
                "Product Team"
            );
        }
    }
}
