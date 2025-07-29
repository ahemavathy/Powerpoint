using PowerPointGenerator;
using PowerPointGenerator.Models;
using PowerPointGenerator.Services;

namespace PowerPointGenerator.Examples
{
    /// <summary>
    /// Complete example showing how to use the updated system with your slide format
    /// </summary>
    public static class YourSlideFormatExample
    {
        /// <summary>
        /// Demonstrates exactly how to use your slide content format
        /// </summary>
        public static async Task RunExample()
        {
            Console.WriteLine("=== Your Slide Format Example ===");

            // Your exact slide content format
            string yourSlideContent = @"
### Slide 1: Introduction
**Title:** Elevating Culinary Experiences  
**Description:** Introducing the next generation of air fryers designed to revolutionize your kitchen experience with cutting-edge technology, elegant design, and unmatched performance that transforms everyday cooking into culinary artistry.
**Suggested Image Background:** Use Image 1: ""air_fryer_rose_gold.jpg"" to immediately convey the premium aspect.

### Slide 2: Sophisticated Design
**Title:** A Design That Speaks Sophistication  
**Description:** Our air fryer, with its sleek contours and premium rose gold finish, seamlessly integrates into modern kitchen aesthetics while delivering exceptional performance that exceeds expectations in every aspect of healthy cooking.
**Suggested Image Background:** Use Image 2: ""kitchen_render.png"" to illustrate its visual appeal in a modern setting.

### Slide 3: Advanced Technology
**Title:** Smart Cooking Innovation
**Description:** Featuring precision temperature control, intelligent cooking programs, and app connectivity that ensures perfectly cooked meals every time, making healthy cooking effortless and enjoyable for the entire family.
**Suggested Image Background:** Use Image 3: ""technology_dashboard.jpg"" to showcase the smart features and digital controls.

### Slide 4: Health Benefits
**Title:** Healthier Cooking Made Easy
**Description:** Reduce oil usage by up to 85% while maintaining incredible taste and texture, enabling you to create nutritious meals that don't compromise on flavor or satisfaction.
**Suggested Image Background:** Use Image 4: ""healthy_food_comparison.jpg"" to show the health benefits visually.
";

            // Method 1: Use the simple API
            await Method1_SimpleAPI(yourSlideContent);

            // Method 2: Use the parser directly
            await Method2_DirectParser(yourSlideContent);

            // Method 3: Create programmatically
            await Method3_Programmatic();

            Console.WriteLine("\n✅ All example presentations created successfully!");
        }

        /// <summary>
        /// Method 1: Using the simple PowerPointAPI
        /// </summary>
        static async Task Method1_SimpleAPI(string slideContent)
        {
            Console.WriteLine("\n🔄 Method 1: Using PowerPointAPI...");

            var outputPath = await PowerPointAPI.CreatePresentationFromText(
                slideContent,
                Path.Combine(Environment.CurrentDirectory, "Method1_AirFryer_Presentation.pptx"),
                "Premium Air Fryer - Product Launch",
                "Marketing Team",
                Path.Combine(Environment.CurrentDirectory, "Images")
            );

            Console.WriteLine($"✅ Created: {Path.GetFileName(outputPath)}");
        }

        /// <summary>
        /// Method 2: Using the SlideContentParser directly
        /// </summary>
        static async Task Method2_DirectParser(string slideContent)
        {
            Console.WriteLine("\n🔄 Method 2: Using SlideContentParser directly...");

            // Parse the content
            var presentationContent = SlideContentParser.ParseSlideContent(
                slideContent,
                "Air Fryer Innovation Showcase",
                "Product Development Team",
                Path.Combine(Environment.CurrentDirectory, "Images")
            );

            // You can modify the parsed content here if needed
            foreach (var slide in presentationContent.Slides)
            {
                // Example: Add bullet points based on slide title
                if (slide.Title.Contains("Technology"))
                {
                    slide.BulletPoints.AddRange(new[]
                    {
                        "Precision temperature control",
                        "Smart cooking programs", 
                        "Mobile app connectivity",
                        "Energy efficient operation"
                    });
                }
            }

            // Generate presentation
            var outputPath = Path.Combine(Environment.CurrentDirectory, "Method2_AirFryer_Enhanced.pptx");
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(presentationContent, outputPath);

            Console.WriteLine($"✅ Created: {Path.GetFileName(outputPath)}");
        }

        /// <summary>
        /// Method 3: Creating content programmatically
        /// </summary>
        static async Task Method3_Programmatic()
        {
            Console.WriteLine("\n🔄 Method 3: Creating programmatically...");

            var content = new PresentationContent
            {
                Title = "Air Fryer - Comprehensive Overview",
                Author = "AI Content Generator",
                Slides = new List<SlideContent>()
            };

            // Slide 1: Introduction with background image
            content.Slides.Add(new SlideContent
            {
                Title = "Elevating Culinary Experiences",
                Description = "Introducing the next generation of air fryers designed to transform your kitchen experience with innovative technology and elegant design.",
                LayoutType = SlideLayoutType.ImageFocused,
                BackgroundImage = new ImageContent
                {
                    FilePath = Path.Combine(Environment.CurrentDirectory, "Images", "air_fryer_rose_gold.jpg"),
                    AltText = "Premium rose gold air fryer",
                    Caption = "Premium design meets functionality"
                },
                BulletPoints = new List<string>
                {
                    "Next-generation technology",
                    "Elegant design aesthetics",
                    "Superior performance"
                }
            });

            // Slide 2: Design focus
            content.Slides.Add(new SlideContent
            {
                Title = "A Design That Speaks Sophistication",
                Description = "Our air fryer, with its sleek contours and premium finish, seamlessly integrates into modern kitchen aesthetics while delivering exceptional performance.",
                LayoutType = SlideLayoutType.ImageFocused,
                BackgroundImage = new ImageContent
                {
                    FilePath = Path.Combine(Environment.CurrentDirectory, "Images", "kitchen_render.png"),
                    AltText = "Modern kitchen with air fryer",
                    Caption = "Perfect integration into modern kitchens"
                }
            });

            // Add the background images to the Images collection for processing
            foreach (var slide in content.Slides.Where(s => s.BackgroundImage != null))
            {
                slide.Images.Add(slide.BackgroundImage!);
            }

            var outputPath = Path.Combine(Environment.CurrentDirectory, "Method3_AirFryer_Programmatic.pptx");
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(content, outputPath);

            Console.WriteLine($"✅ Created: {Path.GetFileName(outputPath)}");
        }

        /// <summary>
        /// Quick test method that you can call from Program.cs
        /// </summary>
        public static async Task QuickTest()
        {
            string testContent = @"
### Slide 1: Introduction
**Title:** Elevating Culinary Experiences  
**Description:** Introducing the next generation of air fryers...  
**Suggested Image Background:** Use Image 1: ""air_fryer_rose_gold.jpg"" to immediately convey the premium aspect.

### Slide 2: Sophisticated Design
**Title:** A Design That Speaks Sophistication  
**Description:** Our air fryer, with its sleek contours...  
**Suggested Image Background:** Use Image 2: ""kitchen_render.png"" to illustrate its visual appeal in a modern setting.
";

            var outputPath = await PowerPointAPI.CreatePresentationFromText(
                testContent,
                "QuickTest_Presentation.pptx",
                "Quick Test Presentation"
            );

            Console.WriteLine($"Quick test presentation created: {outputPath}");
        }
    }
}

// Helper class for integration with external systems
namespace PowerPointGenerator.Integration
{
    /// <summary>
    /// Helper for integrating with external content management systems or AI services
    /// </summary>
    public static class ContentProcessor
    {
        /// <summary>
        /// Processes content from external source (e.g., API response, file, database)
        /// </summary>
        public static async Task<string> ProcessExternalContent(string rawContent, string outputFileName = "External_Content_Presentation.pptx")
        {
            // Clean and format the content if needed
            var cleanedContent = CleanContent(rawContent);

            // Create presentation
            var outputPath = Path.Combine(Environment.CurrentDirectory, outputFileName);
            return await PowerPointAPI.CreatePresentationFromText(
                cleanedContent,
                outputPath,
                "Generated from External Content",
                "Content Processor"
            );
        }

        /// <summary>
        /// Cleans and formats raw content
        /// </summary>
        private static string CleanContent(string rawContent)
        {
            // Add any content cleaning logic here
            // For example: remove extra whitespace, fix formatting, etc.
            return rawContent.Trim();
        }

        /// <summary>
        /// Batch process multiple slide contents
        /// </summary>
        public static async Task<List<string>> ProcessBatch(Dictionary<string, string> contentBatch)
        {
            var results = new List<string>();

            foreach (var kvp in contentBatch)
            {
                var fileName = $"Batch_{kvp.Key}_Presentation.pptx";
                var outputPath = await ProcessExternalContent(kvp.Value, fileName);
                results.Add(outputPath);
            }

            return results;
        }
    }
}
