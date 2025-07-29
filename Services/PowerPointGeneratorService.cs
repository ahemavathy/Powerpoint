using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointGenerator.Models;
using PowerPointGenerator.Utilities;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator.Services
{
    /// <summary>
    /// Service for creating PowerPoint presentations using OpenXML
    /// </summary>
    public class PowerPointGeneratorService : IDisposable
    {
        private PresentationDocument? _presentationDocument;
        private bool _disposed = false;

        /// <summary>
        /// Creates a new PowerPoint presentation from the provided content
        /// </summary>
        /// <param name="content">The presentation content</param>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        /// <returns>Task representing the async operation</returns>
        public async Task CreatePresentationAsync(PresentationContent content, string outputPath)
        {
            try
            {
                // Create a new presentation document
                _presentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);

                // Create the presentation parts
                CreatePresentationParts();

                // Set presentation properties
                SetPresentationProperties(content.Title, content.Author);

                // Create slides
                await CreateSlidesAsync(content.Slides);

                // Save the presentation
                _presentationDocument.Save();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create PowerPoint presentation: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Creates the basic presentation structure
        /// </summary>
        private void CreatePresentationParts()
        {
            if (_presentationDocument == null)
                throw new InvalidOperationException("Presentation document is not initialized");

            // Create the main presentation part
            var presentationPart = _presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            // Create slide master part
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

            // Create slide layout part
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

            // Create theme part
            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = ThemeHelper.CreateDefaultTheme();

            // Initialize presentation structure properly
            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });

            // Add slide size (required for valid presentation)
            presentationPart.Presentation.SlideSize = new SlideSize() 
            { 
                Cx = 9144000, // 10 inches
                Cy = 6858000  // 7.5 inches
            };

            // Add default view properties to prevent repair errors
            presentationPart.Presentation.DefaultTextStyle = new DefaultTextStyle(
                new A.DefaultParagraphProperties(),
                new A.Level1ParagraphProperties(),
                new A.Level2ParagraphProperties(),
                new A.Level3ParagraphProperties(),
                new A.Level4ParagraphProperties(),
                new A.Level5ParagraphProperties(),
                new A.Level6ParagraphProperties(),
                new A.Level7ParagraphProperties(),
                new A.Level8ParagraphProperties(),
                new A.Level9ParagraphProperties());
        }

        /// <summary>
        /// Sets basic presentation properties
        /// </summary>
        private void SetPresentationProperties(string title, string author)
        {
            if (_presentationDocument?.PresentationPart?.Presentation == null)
                return;

            // Set core document properties
            var corePropertiesPart = _presentationDocument.AddCoreFilePropertiesPart();
            using (var writer = new System.Xml.XmlTextWriter(corePropertiesPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                writer.WriteElementString("dc", "title", "http://purl.org/dc/elements/1.1/", title);
                writer.WriteElementString("dc", "creator", "http://purl.org/dc/elements/1.1/", author);
                writer.WriteElementString("dcterms", "created", "http://purl.org/dc/terms/", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteElementString("dcterms", "modified", "http://purl.org/dc/terms/", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        /// <summary>
        /// Creates slides from the provided slide content
        /// </summary>
        private async Task CreateSlidesAsync(List<SlideContent> slides)
        {
            if (_presentationDocument?.PresentationPart == null)
                throw new InvalidOperationException("Presentation part is not initialized");

            var presentationPart = _presentationDocument.PresentationPart;
            var slideIdList = presentationPart.Presentation.SlideIdList;

            uint slideId = 256;

            foreach (var slideContent in slides)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                
                // Create proper slide structure
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties() { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new A.TransformGroup()))),
                    new ColorMapOverride(new A.MasterColorMapping()));

                // Create slide content based on layout type
                await CreateSlideContentAsync(slidePart, slideContent);

                // Add slide to presentation
                var slideIdEntry = new SlideId()
                {
                    Id = slideId++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                slideIdList?.Append(slideIdEntry);
            }
        }

        /// <summary>
        /// Creates content for a specific slide
        /// </summary>
        private Task CreateSlideContentAsync(SlidePart slidePart, SlideContent content)
        {
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree ?? new ShapeTree();
            if (slidePart.Slide.CommonSlideData == null)
            {
                slidePart.Slide.CommonSlideData = new CommonSlideData();
            }
            slidePart.Slide.CommonSlideData.ShapeTree = shapeTree;

            // Add non-visual group shape properties (required)
            if (shapeTree.NonVisualGroupShapeProperties == null)
            {
                shapeTree.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties());
            }

            // Add group shape properties (required)
            if (shapeTree.GroupShapeProperties == null)
            {
                shapeTree.GroupShapeProperties = new GroupShapeProperties(
                    new A.TransformGroup());
            }

            uint shapeId = 2;
            long currentY = 685800; // Start position for title

            // Add title at the top
            if (!string.IsNullOrEmpty(content.Title))
            {
                var titleShape = CreateTitleTextShape(shapeId++, content.Title, currentY);
                shapeTree.Append(titleShape);
                currentY += 1200000; // Move down for next element
            }

            // Add description below title
            if (!string.IsNullOrEmpty(content.Description))
            {
                var descriptionShape = CreateDescriptionTextShape(shapeId++, content.Description, currentY);
                shapeTree.Append(descriptionShape);
                currentY += 1500000; // Move down for images
            }
            // Fallback to synopsis if description is not available
            else if (!string.IsNullOrEmpty(content.Synopsis))
            {
                var synopsisShape = CreateDescriptionTextShape(shapeId++, content.Synopsis, currentY);
                shapeTree.Append(synopsisShape);
                currentY += 1500000; // Move down for images
            }

            // Add images below description/synopsis
            if (content.Images.Any())
            {
                shapeId = AddImagesAtPosition(slidePart, shapeTree, content.Images, shapeId, currentY);
            }

            // Handle background image if it exists and isn't already in Images
            if (content.BackgroundImage != null && !content.Images.Contains(content.BackgroundImage))
            {
                var backgroundImages = new List<ImageContent> { content.BackgroundImage };
                AddImagesAtPosition(slidePart, shapeTree, backgroundImages, shapeId, currentY);
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Creates a title text shape with proper formatting
        /// </summary>
        private Shape CreateTitleTextShape(uint shapeId, string title, long yPosition)
        {
            return SlideHelper.CreateFormattedTextShape(shapeId, title, 
                914400, yPosition, 8229600, 1143000, // Position and size
                2800, true); // Font size 28pt, bold
        }

        /// <summary>
        /// Creates a description text shape with proper formatting
        /// </summary>
        private Shape CreateDescriptionTextShape(uint shapeId, string description, long yPosition)
        {
            return SlideHelper.CreateFormattedTextShape(shapeId, description,
                914400, yPosition, 8229600, 1371600, // Position and size
                1600, false); // Font size 16pt, not bold
        }

        /// <summary>
        /// Adds images to a slide at a specific position
        /// </summary>
        private uint AddImagesAtPosition(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, uint shapeId, long yPosition)
        {
            if (!images.Any()) return shapeId;

            // Calculate image layout
            var slideWidth = 9144000; // Standard slide width in EMUs
            var slideHeight = 6858000; // Standard slide height in EMUs
            var availableHeight = slideHeight - yPosition - 457200; // Leave bottom margin
            
            if (images.Count == 1)
            {
                // Single image - use almost all space below description
                var image = images.First();
                if (File.Exists(image.FilePath))
                {
                    var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                    var minMargin = 100000; // minimal margin
                    var maxWidth = slideWidth - 2 * minMargin;
                    var maxHeight = availableHeight - minMargin;
                    var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
                    var x = (slideWidth - width) / 2;
                    var y = yPosition + minMargin;
                    var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                    using (var stream = File.OpenRead(image.FilePath))
                        imagePart.FeedData(stream);
                    var relationshipId = slidePart.GetIdOfPart(imagePart);
                    var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                    shapeTree.Append(imageShape);
                }
            }
            else
            {
                // Multiple images - arrange in grid, maximize size per cell
                var imagesPerRow = Math.Min(2, images.Count);
                var cellWidth = (slideWidth - 1828800) / imagesPerRow;
                var cellHeight = availableHeight / ((images.Count + imagesPerRow - 1) / imagesPerRow);
                for (int i = 0; i < images.Count; i++)
                {
                    var image = images[i];
                    if (File.Exists(image.FilePath))
                    {
                        var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                        var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, cellWidth - 228600, cellHeight - 228600);
                        var row = i / imagesPerRow;
                        var col = i % imagesPerRow;
                        var x = 914400 + col * cellWidth + (cellWidth - width) / 2;
                        var y = yPosition + row * cellHeight + (cellHeight - height) / 2;
                        var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                        using (var stream = File.OpenRead(image.FilePath))
                            imagePart.FeedData(stream);
                        var relationshipId = slidePart.GetIdOfPart(imagePart);
                        var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                        shapeTree.Append(imageShape);
                    }
                }
            }

            return shapeId;
        }

        /// <summary>
        /// Adds images to a slide based on the layout type
        /// </summary>
        private async Task<uint> AddImagesAsync(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, SlideLayoutType layoutType, uint shapeId)
        {
            if (!images.Any()) return shapeId;

            switch (layoutType)
            {
                case SlideLayoutType.ImageFocused:
                    return await AddSingleLargeImageAsync(slidePart, shapeTree, images.First(), shapeId);

                case SlideLayoutType.ImageGrid:
                    return await AddImageGridAsync(slidePart, shapeTree, images, shapeId);

                case SlideLayoutType.TwoImageComparison:
                    return await AddTwoImageComparisonAsync(slidePart, shapeTree, images.Take(2).ToList(), shapeId);

                case SlideLayoutType.SingleImageWithCaption:
                    return await AddSingleImageWithCaptionAsync(slidePart, shapeTree, images.First(), shapeId);

                default:
                    // Default: add images in a simple layout
                    return await AddDefaultImageLayoutAsync(slidePart, shapeTree, images, shapeId);
            }
        }

        /// <summary>
        /// Adds a single large image to the slide
        /// </summary>
        private Task<uint> AddSingleLargeImageAsync(SlidePart slidePart, ShapeTree shapeTree, 
            ImageContent image, uint shapeId)
        {
            if (!File.Exists(image.FilePath)) return Task.FromResult(shapeId);

            var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
            var maxWidth = 5715000; // About 6.25 inches
            var maxHeight = 4286000; // About 4.7 inches
            var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
            var x = 2286000 + (maxWidth - width) / 2;
            var y = 1524000 + (maxHeight - height) / 2;
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            using (var stream = File.OpenRead(image.FilePath))
                imagePart.FeedData(stream);
            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
            shapeTree.Append(imageShape);
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds images in a grid layout
        /// </summary>
        private Task<uint> AddImageGridAsync(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, uint shapeId)
        {
            const int maxImagesPerRow = 2;
            const long imageWidth = 3600000; // 4 inches
            const long imageHeight = 2700000; // 3 inches
            const long startX = 914400;
            const long startY = 2286000;
            const long spacingX = 4500000;
            const long spacingY = 3200000;

            for (int i = 0; i < images.Count && i < 4; i++) // Max 4 images in grid
            {
                var image = images[i];
                if (!File.Exists(image.FilePath)) continue;

                var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, imageWidth, imageHeight);
                int row = i / maxImagesPerRow;
                int col = i % maxImagesPerRow;
                long x = startX + (col * spacingX) + (imageWidth - width) / 2;
                long y = startY + (row * spacingY) + (imageHeight - height) / 2;
                var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                using (var stream = File.OpenRead(image.FilePath))
                    imagePart.FeedData(stream);
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                shapeTree.Append(imageShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds two images side by side for comparison
        /// </summary>
        private Task<uint> AddTwoImageComparisonAsync(SlidePart slidePart, ShapeTree shapeTree,
            List<ImageContent> images, uint shapeId)
        {
            if (images.Count < 2) return Task.FromResult(shapeId);

            const long imageWidth = 3600000;
            const long imageHeight = 3600000;
            const long leftX = 914400;
            const long rightX = 4914400;
            const long y = 2286000;

            for (int i = 0; i < 2; i++)
            {
                var image = images[i];
                if (!File.Exists(image.FilePath)) continue;

                var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, imageWidth, imageHeight);
                long x = (i == 0 ? leftX : rightX) + (imageWidth - width) / 2;
                long yPos = y + (imageHeight - height) / 2;
                var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                using (var stream = File.OpenRead(image.FilePath))
                    imagePart.FeedData(stream);
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, yPos, width, height, image.AltText);
                shapeTree.Append(imageShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds a single image with caption
        /// </summary>
        private Task<uint> AddSingleImageWithCaptionAsync(SlidePart slidePart, ShapeTree shapeTree,
            ImageContent image, uint shapeId)
        {
            if (!File.Exists(image.FilePath)) return Task.FromResult(shapeId);

            var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
            var maxWidth = 5715000; // About 6.25 inches
            var maxHeight = 4286000; // About 4.7 inches
            var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
            var x = 1828800 + (maxWidth - width) / 2;
            var y = 1524000 + (maxHeight - height) / 2;
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            using (var stream = File.OpenRead(image.FilePath))
                imagePart.FeedData(stream);
            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
            shapeTree.Append(imageShape);

            // Add caption below image
            if (!string.IsNullOrEmpty(image.Caption))
            {
                var captionShape = SlideHelper.CreateTextShape(shapeId++, image.Caption,
                    1828800, 6000000, 5715000, 600000); // Caption below image
                shapeTree.Append(captionShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds images in a default layout
        /// </summary>
        private async Task<uint> AddDefaultImageLayoutAsync(SlidePart slidePart, ShapeTree shapeTree,
            List<ImageContent> images, uint shapeId)
        {
            // For now, use single large image layout for the first image
            if (images.Any())
            {
                return await AddSingleLargeImageAsync(slidePart, shapeTree, images.First(), shapeId);
            }
            return shapeId;
        }

        /// <summary>
        /// Creates an image shape with proper embedding
        /// </summary>
        private Picture CreateImageShape(SlidePart slidePart, ImageContent image, 
            uint shapeId, long x, long y, long width, long height)
        {
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            
            using (var stream = File.OpenRead(image.FilePath))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = slidePart.GetIdOfPart(imagePart);
            
            return SlideHelper.CreateImageShape(shapeId, relationshipId, x, y, width, height, image.AltText);
        }

        /// <summary>
        /// Disposes the presentation document
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected dispose method
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _presentationDocument?.Dispose();
                _disposed = true;
            }
        }
    }
}
