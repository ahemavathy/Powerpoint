using PowerPointGenerator.Models;
using System.Text.RegularExpressions;

namespace PowerPointGenerator.Services
{
    /// <summary>
    /// Parser for structured slide content input format
    /// </summary>
    public class SlideContentParser
    {
        /// <summary>
        /// Parses structured slide content from text format
        /// </summary>
        /// <param name="slideContentText">Multi-slide text content</param>
        /// <param name="presentationTitle">Title for the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="baseImagePath">Base path where images are located</param>
        /// <returns>Parsed presentation content</returns>
        public static PresentationContent ParseSlideContent(string slideContentText, 
            string presentationTitle = "AI Generated Presentation", 
            string author = "AI Assistant",
            string baseImagePath = "")
        {
            var presentation = new PresentationContent
            {
                Title = presentationTitle,
                Author = author,
                Slides = new List<SlideContent>()
            };

            // Split content into individual slides
            var slideBlocks = SplitIntoSlideBlocks(slideContentText);

            foreach (var slideBlock in slideBlocks)
            {
                var slide = ParseSingleSlide(slideBlock, baseImagePath);
                if (slide != null)
                {
                    presentation.Slides.Add(slide);
                }
            }

            return presentation;
        }

        /// <summary>
        /// Splits the input text into individual slide blocks
        /// </summary>
        private static List<string> SplitIntoSlideBlocks(string content)
        {
            var slides = new List<string>();
            
            // Split by slide headers (### Slide X:)
            var slidePattern = @"###\s*Slide\s*\d+\s*:.*?(?=###\s*Slide\s*\d+\s*:|$)";
            var matches = Regex.Matches(content, slidePattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                slides.Add(match.Value.Trim());
            }

            return slides;
        }

        /// <summary>
        /// Parses a single slide block into SlideContent
        /// </summary>
        private static SlideContent? ParseSingleSlide(string slideBlock, string baseImagePath)
        {
            try
            {
                var slide = new SlideContent();

                // Extract title
                var titleMatch = Regex.Match(slideBlock, @"\*\*Title:\*\*\s*(.+?)(?:\n|\r|$)", RegexOptions.Multiline);
                if (titleMatch.Success)
                {
                    slide.Title = titleMatch.Groups[1].Value.Trim();
                }

                // Extract description/synopsis
                var descriptionMatch = Regex.Match(slideBlock, @"\*\*Description:\*\*\s*(.+?)(?:\n\*\*|$)", RegexOptions.Singleline);
                if (descriptionMatch.Success)
                {
                    slide.Description = descriptionMatch.Groups[1].Value.Trim();
                }

                // Extract suggested image background
                var imageMatch = Regex.Match(slideBlock, @"\*\*Suggested Image Background:\*\*\s*(.+?)(?:\n|$)", RegexOptions.Multiline);
                if (imageMatch.Success)
                {
                    var imageInfo = ParseImageInfo(imageMatch.Groups[1].Value.Trim(), baseImagePath);
                    if (imageInfo != null)
                    {
                        slide.BackgroundImage = imageInfo;
                        // Also add to Images collection for processing
                        slide.Images.Add(imageInfo);
                    }
                }

                // Determine layout type based on content
                slide.LayoutType = DetermineLayoutType(slide);

                return slide;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing slide: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Parses image information from the suggested image background text
        /// </summary>
        private static ImageContent? ParseImageInfo(string imageText, string baseImagePath)
        {
            try
            {
                // Extract image filename using regex
                var imageFileMatch = Regex.Match(imageText, @"[""']?([^""'\s:]+\.(jpg|jpeg|png|gif|bmp|tiff))[""']?", RegexOptions.IgnoreCase);
                
                if (imageFileMatch.Success)
                {
                    var fileName = imageFileMatch.Groups[1].Value;
                    var fullPath = string.IsNullOrEmpty(baseImagePath) 
                        ? Path.Combine(Environment.CurrentDirectory, "Images", fileName)
                        : Path.Combine(baseImagePath, fileName);

                    // Extract description/context from the image text
                    var description = ExtractImageDescription(imageText);

                    return new ImageContent
                    {
                        FilePath = fullPath,
                        AltText = description,
                        Caption = description
                    };
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing image info: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Extracts descriptive text about the image from the suggestion
        /// </summary>
        private static string ExtractImageDescription(string imageText)
        {
            // Look for text after "to" which usually contains the purpose/description
            var descriptionMatch = Regex.Match(imageText, @"\s+to\s+(.+)$", RegexOptions.IgnoreCase);
            if (descriptionMatch.Success)
            {
                return descriptionMatch.Groups[1].Value.Trim().TrimEnd('.');
            }

            // Fallback: return a generic description
            return "Background image for slide";
        }

        /// <summary>
        /// Determines the best layout type based on slide content
        /// </summary>
        private static SlideLayoutType DetermineLayoutType(SlideContent slide)
        {
            // If it has background image, make it image-focused
            if (slide.BackgroundImage != null)
            {
                return SlideLayoutType.ImageFocused;
            }

            // If it has multiple images, use grid
            if (slide.Images.Count > 1)
            {
                return SlideLayoutType.ImageGrid;
            }

            // If it has one image, use single image with caption
            if (slide.Images.Count == 1)
            {
                return SlideLayoutType.SingleImageWithCaption;
            }

            // Default to title and content
            return SlideLayoutType.TitleAndContent;
        }

        /// <summary>
        /// Parses slide content from a structured format with more flexibility
        /// </summary>
        public static PresentationContent ParseFlexibleFormat(string content, 
            string presentationTitle = "AI Generated Presentation",
            string author = "AI Assistant",
            string baseImagePath = "")
        {
            var presentation = new PresentationContent
            {
                Title = presentationTitle,
                Author = author,
                Slides = new List<SlideContent>()
            };

            // Handle different input formats
            if (content.Contains("### Slide"))
            {
                return ParseSlideContent(content, presentationTitle, author, baseImagePath);
            }
            else
            {
                // Handle other formats or create a single slide
                var slide = new SlideContent
                {
                    Title = "Generated Content",
                    Description = content,
                    LayoutType = SlideLayoutType.TitleAndContent
                };
                presentation.Slides.Add(slide);
            }

            return presentation;
        }

        /// <summary>
        /// Creates a sample slide content for testing
        /// </summary>
        public static string CreateSampleSlideContent()
        {
            return @"
### Slide 1: Introduction
**Title:** Elevating Culinary Experiences
**Description:** Introducing the next generation of air fryers designed to transform your kitchen experience with innovative technology and elegant design.
**Suggested Image Background:** Use Image 1: ""air_fryer_rose_gold.jpg"" to immediately convey the premium aspect.

### Slide 2: Sophisticated Design
**Title:** A Design That Speaks Sophistication
**Description:** Our air fryer, with its sleek contours and premium finish, seamlessly integrates into modern kitchen aesthetics while delivering exceptional performance.
**Suggested Image Background:** Use Image 2: ""kitchen_render.png"" to illustrate its visual appeal in a modern setting.

### Slide 3: Advanced Technology
**Title:** Innovation Meets Performance
**Description:** Featuring cutting-edge heating technology and precision controls that ensure perfectly cooked meals every time, making healthy cooking effortless and enjoyable.
**Suggested Image Background:** Use Image 3: ""technology_closeup.jpg"" to showcase the advanced features and controls.
";
        }
    }
}
