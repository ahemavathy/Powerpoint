using PowerPointGenerator.Models;

namespace PowerPointGenerator.Examples
{
    /// <summary>
    /// Example class showing how to integrate AI-generated content
    /// </summary>
    public class AIContentIntegration
    {
        /// <summary>
        /// Example method showing how to structure AI-generated content for presentation creation
        /// </summary>
        /// <param name="aiGeneratedData">Raw AI content</param>
        /// <param name="imagePaths">Paths to AI-generated or selected images</param>
        /// <returns>Structured presentation content</returns>
        public static PresentationContent ProcessAIContent(AIGeneratedData aiGeneratedData, List<string> imagePaths)
        {
            var presentation = new PresentationContent
            {
                Title = aiGeneratedData.Title,
                Author = aiGeneratedData.Author ?? "AI Assistant",
                Slides = new List<SlideContent>()
            };

            // Process each AI-generated section into slides
            for (int i = 0; i < aiGeneratedData.Sections.Count; i++)
            {
                var section = aiGeneratedData.Sections[i];
                var slide = new SlideContent
                {
                    Title = section.Title,
                    Synopsis = section.Synopsis,
                    BulletPoints = section.KeyPoints,
                    LayoutType = DetermineLayoutType(section, imagePaths.Count),
                    Images = AssignImagesToSlide(section, imagePaths, i)
                };

                presentation.Slides.Add(slide);
            }

            return presentation;
        }

        /// <summary>
        /// Determines the best layout type based on content and available images
        /// </summary>
        private static SlideLayoutType DetermineLayoutType(AISection section, int totalImages)
        {
            // Logic to determine layout based on content characteristics
            if (section.IsComparison)
                return SlideLayoutType.TwoImageComparison;
            
            if (section.RequiresMultipleVisuals && totalImages >= 4)
                return SlideLayoutType.ImageGrid;
            
            if (!string.IsNullOrEmpty(section.DetailedCaption))
                return SlideLayoutType.SingleImageWithCaption;
            
            if (section.IsVisuallyFocused)
                return SlideLayoutType.ImageFocused;

            return SlideLayoutType.TitleAndContent;
        }

        /// <summary>
        /// Assigns appropriate images to each slide based on content
        /// </summary>
        private static List<ImageContent> AssignImagesToSlide(AISection section, List<string> imagePaths, int slideIndex)
        {
            var images = new List<ImageContent>();

            // Simple assignment strategy - can be enhanced with AI-based image-text matching
            var imagesPerSlide = Math.Min(section.RequestedImageCount, imagePaths.Count - slideIndex);
            
            for (int i = 0; i < imagesPerSlide; i++)
            {
                var imageIndex = (slideIndex * 2 + i) % imagePaths.Count; // Distribute images across slides
                
                if (imageIndex < imagePaths.Count)
                {
                    images.Add(new ImageContent
                    {
                        FilePath = imagePaths[imageIndex],
                        AltText = section.ImageDescriptions.ElementAtOrDefault(i) ?? "AI-generated visual",
                        Caption = section.ImageCaptions.ElementAtOrDefault(i) ?? ""
                    });
                }
            }

            return images;
        }
    }

    /// <summary>
    /// Represents AI-generated content structure
    /// </summary>
    public class AIGeneratedData
    {
        public string Title { get; set; } = string.Empty;
        public string? Author { get; set; }
        public List<AISection> Sections { get; set; } = new List<AISection>();
    }

    /// <summary>
    /// Represents a section of AI-generated content
    /// </summary>
    public class AISection
    {
        public string Title { get; set; } = string.Empty;
        public string Synopsis { get; set; } = string.Empty;
        public List<string> KeyPoints { get; set; } = new List<string>();
        public List<string> ImageDescriptions { get; set; } = new List<string>();
        public List<string> ImageCaptions { get; set; } = new List<string>();
        public int RequestedImageCount { get; set; } = 1;
        public bool IsComparison { get; set; }
        public bool RequiresMultipleVisuals { get; set; }
        public bool IsVisuallyFocused { get; set; }
        public string DetailedCaption { get; set; } = string.Empty;
    }

    /// <summary>
    /// Example usage class
    /// </summary>
    public static class UsageExample
    {
        /// <summary>
        /// Demonstrates how to use the AI content integration
        /// </summary>
        public static async Task<string> CreatePresentationFromAIContent()
        {
            // Example: Simulate AI-generated content
            var aiData = new AIGeneratedData
            {
                Title = "Market Analysis Report",
                Author = "AI Business Analyst",
                Sections = new List<AISection>
                {
                    new AISection
                    {
                        Title = "Market Overview",
                        Synopsis = "Current market conditions show significant growth opportunities in emerging sectors.",
                        KeyPoints = new List<string> { "25% growth in Q3", "Emerging markets leading", "Technology adoption rising" },
                        RequestedImageCount = 1,
                        IsVisuallyFocused = true,
                        ImageDescriptions = new List<string> { "Market growth chart" }
                    },
                    new AISection
                    {
                        Title = "Competitive Landscape",
                        Synopsis = "Analysis of key competitors reveals strategic opportunities for market positioning.",
                        KeyPoints = new List<string> { "3 major competitors identified", "Pricing gaps discovered", "Innovation opportunities" },
                        RequestedImageCount = 4,
                        RequiresMultipleVisuals = true,
                        ImageDescriptions = new List<string> { "Competitor A overview", "Competitor B analysis", "Market share diagram", "Pricing comparison" }
                    },
                    new AISection
                    {
                        Title = "Before vs After Implementation",
                        Synopsis = "Comparison of performance metrics before and after strategy implementation.",
                        IsComparison = true,
                        RequestedImageCount = 2,
                        ImageDescriptions = new List<string> { "Before metrics", "After metrics" },
                        ImageCaptions = new List<string> { "Previous performance baseline", "Post-implementation results" }
                    }
                }
            };

            // Example image paths (replace with actual AI-generated or selected images)
            var imagePaths = new List<string>
            {
                @"C:\Images\chart1.png",
                @"C:\Images\chart2.png", 
                @"C:\Images\diagram1.png",
                @"C:\Images\comparison_before.png",
                @"C:\Images\comparison_after.png"
            };

            // Process AI content into presentation format
            var presentationContent = AIContentIntegration.ProcessAIContent(aiData, imagePaths);

            // Generate the presentation
            var outputPath = Path.Combine(Environment.CurrentDirectory, "AI_Processed_Presentation.pptx");
            
            using var generator = new PowerPointGenerator.Services.PowerPointGeneratorService();
            await generator.CreatePresentationAsync(presentationContent, outputPath);

            return outputPath;
        }
    }
}
