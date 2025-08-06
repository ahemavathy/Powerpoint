# PowerPoint Generator Console Application

A C# .NET 8.0 console application that generates PowerPoint presentations from structured JSON content using the Open-XML-SDK.

## Features

- **JSON-Based Input**: Create presentations from structured JSON files
- **Image Integration**: Automatically includes images with proper aspect ratio preservation
- **Clean Text Formatting**: Titles and descriptions without bullet points
- **Placeholder Generation**: Creates placeholder images when originals are missing
- **Command Line Interface**: Easy to use with flexible arguments
- **Multiple Layout Types**: Supports various slide layouts based on content

## Quick Start

### 1. Basic Usage

Create a JSON file with your slide content:

```json
{
  "slides": [
    {
      "title": "Your Slide Title",
      "description": "Your slide description text goes here.",
      "suggested_image": "your_image.png"
    }
  ]
}
```

Run the console application:
```bash
# From the root directory
dotnet run --project PowerPointGenerator.csproj
```

### 2. Command Line Usage

The program supports flexible command line arguments:

```bash
# Use defaults (slides_content.json → slides_content.pptx)
dotnet run --project PowerPointGenerator.csproj

# Specify JSON file only (uses filename for presentation name)
dotnet run --project PowerPointGenerator.csproj path/to/your/slides.json

# Specify both JSON file and presentation name
dotnet run --project PowerPointGenerator.csproj path/to/your/slides.json "My Presentation Name"

# Examples:
dotnet run --project PowerPointGenerator.csproj slides_content.json "Product_Showcase"
dotnet run --project PowerPointGenerator.csproj data/slides.json "Q4-Sales-Report"
dotnet run --project PowerPointGenerator.csproj content.json "Marketing Presentation"
```

**Command Line Arguments:**
1. **First argument**: JSON file path (optional, defaults to `slides_content.json`)
2. **Second argument**: Presentation name (optional, defaults to JSON filename without extension)

**Behavior:**
- The default presentation name is derived from the JSON filename (e.g., `my_slides.json` → `my_slides.pptx`)
- Use quotes around presentation names that contain spaces
- Underscores and dashes in names are converted to spaces in the presentation title
- The `.pptx` extension is added automatically

## JSON Format

The expected JSON format is:

```json
{
  "slides": [
    {
      "title": "Slide Title",
      "description": "Slide description text",
      "suggested_image": "Use Image 1: image_filename.png",
      "layout": "imagefocused"
    },
    {
      "title": "Product Showcase",
      "description": "Side-by-side layout with title/description on left and image on right",
      "suggested_image": "product.jpg",
      "layout": "productshowcase"
    }
  ]
}
```

### Supported Layout Types

- **`imagefocused`** (default) - Large centered image below title and description
- **`productshowcase`** - Side-by-side layout with text on left (40%) and image on right (60%)
- **`imagegrid`** - Multiple images arranged in a grid
- **`twoimagecomparison`** - Two images side by side for comparison
- **`singleimagewithcaption`** - Single image with detailed caption below
- **`titleandcontent`** - Standard title and content layout
- **`title`** - Title-only slide

### Supported Image Formats

- PNG (.png)
- JPEG (.jpg, .jpeg)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tiff, .tif)
- WebP (.webp)

### Image Path Formats

The `suggested_image` field supports various formats and will extract the filename:
- `"Use Image 1: filename.png"`
- `"Use Image 2: "filename.png""`
- `"filename.png"`
- `""filename.png""`

## Project Structure

```
PowerPointGenerator/
├── Models/
│   ├── PresentationModels.cs      # Core domain models
│   └── JsonSlideModels.cs         # JSON-specific models
├── Services/
│   ├── PowerPointGeneratorService.cs  # Main generation logic
│   ├── JsonSlideParser.cs         # JSON parsing logic
│   └── SlideContentParser.cs      # Legacy text parsing
├── Utilities/
│   ├── SlideHelper.cs             # Slide creation helpers
│   ├── ImageHelper.cs             # Image processing with aspect ratio
│   └── ThemeHelper.cs             # Theme creation helpers
├── Images/                        # Place your images here
├── Program.cs                     # Console application entry point
├── PowerPointAPI.cs               # Simplified API wrapper
├── slides_content.json            # Sample JSON file
└── PowerPointGenerator.csproj     # Project file
```

## Layout Features

- **Responsive Layouts**: Automatically chooses optimal layout based on content and specified layout type
- **Aspect Ratio Preservation**: Images maintain their original proportions
- **Multiple Layout Types**: Choose from 7 different slide layouts
  - **Standard Layouts**: Title at top, description below, images below text
  - **ProductShowcase**: Side-by-side layout with text on left (40%) and large image on right (60%)
- **Smart Text Formatting**: Clean title formatting, bullet point conversion for product descriptions
- **Large Image Display**: Images are scaled to use maximum available space within their layout area

## Image Handling

- **Aspect Ratio Preservation**: Images are scaled while maintaining their original proportions
- **Centered Positioning**: Images are centered within their allocated space
- **Multiple Layout Support**: 
  - **Single large image** - Centered below content
  - **Product showcase** - Large image on right side with text on left
  - **Image grid** - 2x2 grid for multiple images
  - **Image with captions** - Detailed image descriptions
  - **Two-image comparison** - Side-by-side image comparison
  - **Image focused** - Minimal text with large image emphasis

## API Reference

### PowerPointAPI Class

#### CreatePresentationFromJsonFile
```csharp
public static async Task<string> CreatePresentationFromJsonFile(
    string jsonFilePath,
    string outputPath,
    string presentationTitle,
    string author)
```

#### CreatePresentationFromJsonString
```csharp
public static async Task<string> CreatePresentationFromJsonString(
    string jsonContent,
    string outputPath,
    string presentationTitle,
    string author)
```

### JsonSlideParser Class

#### ParseFromFile
```csharp
public static PresentationContent ParseFromFile(
    string jsonFilePath,
    string presentationTitle,
    string author,
    string? imageBasePath = null)
```

#### ParseFromString
```csharp
public static PresentationContent ParseFromString(
    string jsonContent,
    string presentationTitle,
    string author,
    string? imageBasePath = null)
```

## Requirements

- **.NET 8.0 SDK** - Required to build and run the application
- **Windows OS** - For System.Drawing.Common image processing
- **Sufficient disk space** - For generated presentations and images

## File Locations

- **Input**: JSON files can be placed anywhere (specify path as argument)
- **Images**: Place image files in the `Images/` directory in the project root
- **Output**: Generated presentations are saved in the project root directory

## Error Handling

The application includes comprehensive error handling for:
- **Missing JSON files** - Clear error messages with file path
- **Invalid JSON format** - Detailed parsing error information
- **Missing image files** - Creates colored placeholder images automatically
- **OpenXML generation errors** - Comprehensive error reporting
- **File I/O errors** - Permission and disk space issues

## Output Compatibility

Generated PowerPoint files are fully compatible with:
- **Microsoft PowerPoint 2013 and later**
- **PowerPoint Online**
- **LibreOffice Impress**
- **Google Slides** (with import)

## Example Usage

```bash
# Create a presentation from the sample JSON
dotnet run --project PowerPointGenerator.csproj

# This will create slides_content.pptx from slides_content.json
```

Sample output files:
- `slides_content.pptx` - Main presentation file
- Placeholder images in `Images/` folder (if originals are missing)

## Troubleshooting

**Common Issues:**

1. **Build errors**: Ensure .NET 8.0 SDK is installed
   ```bash
   dotnet --version  # Should show 8.0.x
   ```

2. **Missing images**: Check that image files are in the `Images/` directory or update the JSON to point to correct paths

3. **JSON format errors**: Validate your JSON using online tools or check the console error messages

4. **Permission errors**: Ensure write permissions to the output directory

5. **File not found**: Use absolute paths or ensure files are relative to the project directory

## Advanced Usage

### Custom Image Directory
You can specify a custom image directory by modifying the `ParseFromFile` call in your code:

```csharp
var content = JsonSlideParser.ParseFromFile(
    jsonFilePath,
    presentationTitle,
    author,
    @"C:\path\to\your\images"  // Custom image directory
);
```

### Programmatic Usage
For integration into other applications:

```csharp
using PowerPointGenerator.Services;

// Parse JSON and create presentation
var content = JsonSlideParser.ParseFromFile("data.json", "My Presentation", "Author");
using var generator = new PowerPointGeneratorService();
await generator.CreatePresentationAsync(content, "output.pptx");
```

## License

This project is available under the MIT License. See the LICENSE file for more details.

---

**Ready to generate professional PowerPoint presentations from JSON! 🎉**
