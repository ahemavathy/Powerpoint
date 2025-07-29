using PowerPointGenerator.Models;
using PowerPointGenerator.Services;
using PowerPointGenerator.Utilities;
using System.Drawing;

namespace PowerPointGenerator
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                Console.WriteLine("PowerPoint Generator - Creating presentations from JSON slide content");
                Console.WriteLine("=========================================================================");

                // Parse command line arguments
                string jsonFilePath;
                string presentationName;

                // First argument: JSON file path
                if (args.Length > 0 && File.Exists(args[0]))
                {
                    jsonFilePath = args[0];
                    Console.WriteLine($"Using JSON file: {jsonFilePath}");
                }
                else if (args.Length > 0 && !File.Exists(args[0]))
                {
                    Console.WriteLine($"⚠️  Warning: Specified JSON file '{args[0]}' not found. Using default.");
                    jsonFilePath = Path.Combine(Environment.CurrentDirectory, "slides_content.json");
                    Console.WriteLine($"Using default JSON file: {jsonFilePath}");
                }
                else
                {
                    jsonFilePath = Path.Combine(Environment.CurrentDirectory, "slides_content.json");
                    Console.WriteLine($"Using default JSON file: {jsonFilePath}");
                }

                // Second argument: Presentation name
                if (args.Length > 1 && !string.IsNullOrWhiteSpace(args[1]))
                {
                    presentationName = args[1];
                    Console.WriteLine($"Using presentation name: {presentationName}");
                }
                else
                {
                    // Use JSON filename (without extension) as default presentation name
                    presentationName = Path.GetFileNameWithoutExtension(jsonFilePath);
                    Console.WriteLine($"Using presentation name from JSON filename: {presentationName}");
                }

                // Parse JSON and create presentation
                await CreatePresentationFromJsonFile(jsonFilePath, presentationName);

                Console.WriteLine("\n✅ Presentation created successfully!");
                Console.WriteLine($"\nFile created: {presentationName}.pptx");
                
                Console.WriteLine("\nUsage options:");
                Console.WriteLine("1. dotnet run                                    # Use default JSON (name from file)");
                Console.WriteLine("2. dotnet run slides.json                       # Use JSON file (name from file)");
                Console.WriteLine("3. dotnet run slides.json \"My Presentation\"     # Specify JSON file and custom name");
                Console.WriteLine("\nTips:");
                Console.WriteLine("- The presentation name defaults to the JSON filename (without .json extension)");
                Console.WriteLine("- Create a JSON file with your slide content (see slides_content.json for format)");
                Console.WriteLine("- Place your images in the 'Images' folder");
                Console.WriteLine("- Use quotes around presentation names with spaces");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating presentation: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Creates a presentation from JSON slide content
        /// </summary>
        /// <param name="jsonFilePath">Path to the JSON file</param>
        /// <param name="presentationName">Name for the output presentation file (without extension)</param>
        static async Task CreatePresentationFromJsonFile(string jsonFilePath, string presentationName)
        {
            Console.WriteLine($"\n🔄 Processing JSON slide content from: {Path.GetFileName(jsonFilePath)}");

            // Ensure Images directory exists
            var imageDirectory = Path.Combine(Environment.CurrentDirectory, "Images");
            if (!Directory.Exists(imageDirectory))
            {
                Directory.CreateDirectory(imageDirectory);
                Console.WriteLine($"📁 Created Images directory: {imageDirectory}");
            }

            // Extract presentation title from the name (capitalize and format nicely)
            var presentationTitle = FormatPresentationTitle(presentationName);

            // Parse JSON content using the new parser
            var presentationContent = JsonSlideParser.ParseFromFile(
                jsonFilePath,
                presentationTitle,
                "Product Marketing Team",
                imageDirectory
            );

            // Generate the presentation with custom name
            var outputPath = Path.Combine(Environment.CurrentDirectory, $"{presentationName}.pptx");
            
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(presentationContent, outputPath);

            Console.WriteLine($"✅ Presentation created: {outputPath}");
        }

        /// <summary>
        /// Formats a presentation name into a proper title
        /// </summary>
        /// <param name="name">Raw presentation name</param>
        /// <returns>Formatted presentation title</returns>
        static string FormatPresentationTitle(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Generated Presentation";

            // Replace underscores and dashes with spaces
            var formatted = name.Replace('_', ' ').Replace('-', ' ');
            
            // Capitalize first letter of each word
            var words = formatted.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 0)
                {
                    words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
                }
            }
            
            return string.Join(" ", words);
        }

    }
}
