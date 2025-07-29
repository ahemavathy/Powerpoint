using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator.Utilities
{
    /// <summary>
    /// Helper class for creating slide elements
    /// </summary>
    public static class SlideHelper
    {
        /// <summary>
        /// Creates a text shape with the specified content and position
        /// </summary>
        /// <param name="shapeId">Unique shape ID</param>
        /// <param name="text">Text content</param>
        /// <param name="x">X position in EMUs</param>
        /// <param name="y">Y position in EMUs</param>
        /// <param name="width">Width in EMUs</param>
        /// <param name="height">Height in EMUs</param>
        /// <returns>Shape element</returns>
        public static Shape CreateTextShape(uint shapeId, string text, long x, long y, long width, long height)
        {
            var shape = new Shape();

            var nonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = shapeId, Name = $"TextBox {shapeId}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape()));

            var shapeProperties = new ShapeProperties();

            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = width, Cy = height };

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new A.AdjustValueList());

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            var textBody = new TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            var paragraph = new A.Paragraph();
            var run = new A.Run();
            run.Append(new A.Text() { Text = text });
            paragraph.Append(run);

            textBody.Append(paragraph);

            shape.Append(nonVisualShapeProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            return shape;
        }

        /// <summary>
        /// Creates an image shape with the specified properties
        /// </summary>
        /// <param name="shapeId">Unique shape ID</param>
        /// <param name="relationshipId">Relationship ID to the image</param>
        /// <param name="x">X position in EMUs</param>
        /// <param name="y">Y position in EMUs</param>
        /// <param name="width">Width in EMUs</param>
        /// <param name="height">Height in EMUs</param>
        /// <param name="altText">Alternative text for accessibility</param>
        /// <returns>Picture element</returns>
        public static Picture CreateImageShape(uint shapeId, string relationshipId, 
            long x, long y, long width, long height, string altText = "")
        {
            var picture = new Picture();

            var nonVisualPictureProperties = new NonVisualPictureProperties(
                new NonVisualDrawingProperties() 
                { 
                    Id = shapeId, 
                    Name = $"Picture {shapeId}",
                    Description = altText
                },
                new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties());

            var blipFill = new BlipFill();
            var blip = new A.Blip() { Embed = relationshipId };
            var stretch = new A.Stretch(new A.FillRectangle());
            blipFill.Append(blip);
            blipFill.Append(stretch);

            var shapeProperties = new ShapeProperties();

            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = width, Cy = height };

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new A.AdjustValueList());

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            return picture;
        }

        /// <summary>
        /// Creates a title shape with larger font formatting
        /// </summary>
        /// <param name="shapeId">Unique shape ID</param>
        /// <param name="title">Title text</param>
        /// <param name="x">X position in EMUs</param>
        /// <param name="y">Y position in EMUs</param>
        /// <param name="width">Width in EMUs</param>
        /// <param name="height">Height in EMUs</param>
        /// <returns>Shape element</returns>
        public static Shape CreateTitleShape(uint shapeId, string title, long x, long y, long width, long height)
        {
            var shape = new Shape();

            var nonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = shapeId, Name = $"Title {shapeId}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));

            var shapeProperties = new ShapeProperties();

            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = width, Cy = height };

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new A.AdjustValueList());

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            var textBody = new TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            var paragraph = new A.Paragraph();
            var paragraphProperties = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            paragraph.Append(paragraphProperties);

            var run = new A.Run();
            var runProperties = new A.RunProperties() { FontSize = 4400 }; // Larger font size for titles
            run.Append(runProperties);
            run.Append(new A.Text() { Text = title });
            paragraph.Append(run);

            textBody.Append(paragraph);

            shape.Append(nonVisualShapeProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            return shape;
        }

        /// <summary>
        /// Creates a content shape with bullet points
        /// </summary>
        /// <param name="shapeId">Unique shape ID</param>
        /// <param name="content">Content text with bullet points</param>
        /// <param name="x">X position in EMUs</param>
        /// <param name="y">Y position in EMUs</param>
        /// <param name="width">Width in EMUs</param>
        /// <param name="height">Height in EMUs</param>
        /// <returns>Shape element</returns>
        public static Shape CreateBulletShape(uint shapeId, List<string> bulletPoints, long x, long y, long width, long height)
        {
            var shape = new Shape();

            var nonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = shapeId, Name = $"Content {shapeId}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body }));

            var shapeProperties = new ShapeProperties();

            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = width, Cy = height };

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new A.AdjustValueList());

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            var textBody = new TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            // Create paragraphs for each bullet point
            foreach (var bulletPoint in bulletPoints)
            {
                var paragraph = new A.Paragraph();
                var paragraphProperties = new A.ParagraphProperties() { Level = 0 };
                
                // Add bullet formatting
                var bulletFont = new A.BulletFont() { Typeface = "Arial" };
                var bulletChar = new A.CharacterBullet() { Char = "•" };
                paragraphProperties.Append(bulletFont);
                paragraphProperties.Append(bulletChar);
                
                paragraph.Append(paragraphProperties);

                var run = new A.Run();
                run.Append(new A.Text() { Text = bulletPoint });
                paragraph.Append(run);

                textBody.Append(paragraph);
            }

            shape.Append(nonVisualShapeProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            return shape;
        }

        /// <summary>
        /// Creates a formatted text shape with specific font size and style
        /// </summary>
        /// <param name="shapeId">Unique shape ID</param>
        /// <param name="text">Text content</param>
        /// <param name="x">X position in EMUs</param>
        /// <param name="y">Y position in EMUs</param>
        /// <param name="width">Width in EMUs</param>
        /// <param name="height">Height in EMUs</param>
        /// <param name="fontSize">Font size (in hundredths of a point)</param>
        /// <param name="bold">Whether text should be bold</param>
        /// <returns>Shape element</returns>
        public static Shape CreateFormattedTextShape(uint shapeId, string text, long x, long y, long width, long height, int fontSize = 1800, bool bold = false)
        {
            var shape = new Shape();

            // Remove PlaceholderShape to avoid default bullet formatting
            var nonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = shapeId, Name = $"TextBox {shapeId}" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties());

            var shapeProperties = new ShapeProperties();

            var transform2D = new A.Transform2D();
            transform2D.Offset = new A.Offset() { X = x, Y = y };
            transform2D.Extents = new A.Extents() { Cx = width, Cy = height };

            var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            presetGeometry.Append(new A.AdjustValueList());

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            var textBody = new TextBody();
            textBody.Append(new A.BodyProperties());
            textBody.Append(new A.ListStyle());

            var paragraph = new A.Paragraph();
            
            // Explicitly disable bullets by setting paragraph properties
            var paragraphProperties = new A.ParagraphProperties();
            paragraphProperties.Append(new A.NoBullet()); // This removes bullets
            paragraph.Append(paragraphProperties);
            
            var run = new A.Run();
            var runProperties = new A.RunProperties() { FontSize = fontSize };
            if (bold)
            {
                runProperties.Bold = true;
            }
            
            run.Append(runProperties);
            run.Append(new A.Text() { Text = text });
            paragraph.Append(run);

            textBody.Append(paragraph);

            shape.Append(nonVisualShapeProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            return shape;
        }
    }
}
