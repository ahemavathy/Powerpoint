using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator.Utilities
{
    /// <summary>
    /// Helper class for working with images in PowerPoint
    /// </summary>
    public static class ImageHelper
    {
        /// <summary>
        /// Gets the appropriate content type string based on file extension
        /// </summary>
        /// <param name="filePath">Path to the image file</param>
        /// <returns>Content type string for the file</returns>
        public static string GetImageContentType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            return extension switch
            {
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".tiff" or ".tif" => "image/tiff",
                _ => "image/jpeg" // Default to JPEG
            };
        }

        /// <summary>
        /// Creates an ImagePart with the appropriate type
        /// </summary>
        /// <param name="slidePart">The slide part to add the image to</param>
        /// <param name="filePath">Path to the image file</param>
        /// <returns>The created ImagePart</returns>
        public static ImagePart CreateImagePart(SlidePart slidePart, string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            return extension switch
            {
                ".jpg" or ".jpeg" => slidePart.AddImagePart(ImagePartType.Jpeg),
                ".png" => slidePart.AddImagePart(ImagePartType.Png),
                ".gif" => slidePart.AddImagePart(ImagePartType.Gif),
                ".bmp" => slidePart.AddImagePart(ImagePartType.Bmp),
                ".tiff" or ".tif" => slidePart.AddImagePart(ImagePartType.Tiff),
                _ => slidePart.AddImagePart(ImagePartType.Jpeg) // Default to JPEG
            };
        }

        /// <summary>
        /// Gets the dimensions of an image file
        /// </summary>
        /// <param name="filePath">Path to the image file</param>
        /// <returns>Size of the image in pixels</returns>
        public static Size GetImageDimensions(string filePath)
        {
            try
            {
                if (OperatingSystem.IsWindowsVersionAtLeast(6, 1))
                {
                    using var image = Image.FromFile(filePath);
                    return new Size(image.Width, image.Height);
                }
                else
                {
                    // Return default size for unsupported platforms
                    return new Size(800, 600);
                }
            }
            catch
            {
                // Return default size if unable to read image
                return new Size(800, 600);
            }
        }

        /// <summary>
        /// Converts pixels to EMUs (English Metric Units)
        /// </summary>
        /// <param name="pixels">Size in pixels</param>
        /// <param name="dpi">Dots per inch (default: 96)</param>
        /// <returns>Size in EMUs</returns>
        public static long PixelsToEmus(int pixels, double dpi = 96.0)
        {
            // 1 inch = 914400 EMUs
            // EMUs = pixels * 914400 / DPI
            return (long)(pixels * 914400.0 / dpi);
        }

        /// <summary>
        /// Calculates the best fit dimensions for an image within specified bounds
        /// </summary>
        /// <param name="imageSize">Original image size</param>
        /// <param name="maxWidth">Maximum width in EMUs</param>
        /// <param name="maxHeight">Maximum height in EMUs</param>
        /// <returns>Fitted dimensions in EMUs</returns>
        public static (long width, long height) CalculateFitDimensions(Size imageSize, long maxWidth, long maxHeight)
        {
            var imageWidthEmus = PixelsToEmus(imageSize.Width);
            var imageHeightEmus = PixelsToEmus(imageSize.Height);

            // If image fits within bounds, return original size
            if (imageWidthEmus <= maxWidth && imageHeightEmus <= maxHeight)
            {
                return (imageWidthEmus, imageHeightEmus);
            }

            // Calculate scaling ratios
            var widthRatio = (double)maxWidth / imageWidthEmus;
            var heightRatio = (double)maxHeight / imageHeightEmus;

            // Use the smaller ratio to maintain aspect ratio
            var scale = Math.Min(widthRatio, heightRatio);

            return ((long)(imageWidthEmus * scale), (long)(imageHeightEmus * scale));
        }
    }
}
