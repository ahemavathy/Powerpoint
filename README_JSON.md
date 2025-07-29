# PowerPoint Generator

A C# .NET 8.0 application that generates PowerPoint presentations from structured JSON content using the Open-XML-SDK.

## Features

- **JSON-Based Input**: Create presentations from structured JSON files
- **Image Integration**: Automatically includes images with proper layout
- **Clean Text Formatting**: Titles and descriptions without bullet points
- **Placeholder Generation**: Creates placeholder images when originals are missing
- **Flexible API**: Multiple ways to create presentations

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

Run the program:
```bash
dotnet run
```

### 2. Command Line Usage

The program supports flexible command line arguments:

```bash
# Use defaults (slides_content.json + "Generated_Presentation.pptx")
dotnet run

# Specify JSON file only
dotnet run path/to/your/slides.json

# Specify both JSON file and presentation name
dotnet run path/to/your/slides.json "My Presentation Name"

# Examples:
dotnet run slides_content.json "Product_Showcase"
dotnet run data/slides.json "Q4-Sales-Report"
dotnet run content.json "Marketing Presentation"
```

**Command Line Arguments:**
1. **First argument**: JSON file path (optional, defaults to `slides_content.json`)
2. **Second argument**: Presentation name (optional, defaults to JSON filename without extension)

**Notes:**
- The default presentation name is derived from the JSON filename (e.g., `my_slides.json` → `my_slides.pptx`)
- Use quotes around presentation names that contain spaces
- Underscores and dashes in names will be converted to spaces in the presentation title
- The `.pptx` extension is added automatically

### 3. Programmatic Usage

```csharp
// From JSON file
var outputPath = await PowerPointAPI.CreatePresentationFromJsonFile(
    "slides_content.json",
    "output.pptx",
    "My Presentation",
    "Author Name"
);

// From JSON string
var jsonContent = """{"slides": [...]}""";
var outputPath = await PowerPointAPI.CreatePresentationFromJsonString(
    jsonContent,
    "output.pptx"
);
```

## JSON Format

The expected JSON format is:

```json
{
  "slides": [
    {
      "title": "Slide Title",
      "description": "Slide description text",
      "suggested_image": "Use Image 1: image_filename.png"
    }
  ]
}
```

### Supported Image Formats

- PNG (.png)
- JPEG (.jpg, .jpeg)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tiff, .tif)

### Image Path Formats

The `suggested_image` field supports various formats:
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
│   └── SlideContentParser.cs      # Legacy text parsing (still supported)
├── Utilities/
│   ├── SlideHelper.cs             # Slide creation helpers
│   ├── ImageHelper.cs             # Image processing helpers
│   └── ThemeHelper.cs             # Theme creation helpers
├── Examples/
│   └── JsonExample.cs             # Usage examples
└── slides_content.json            # Sample JSON file
```

## Layout Features

- **Title at Top**: Clean title formatting without bullets
- **Description Below Title**: Properly formatted description text
- **Images Below Text**: Centered image placement with proper sizing
- **Automatic Layout**: Responsive layout based on content

## Image Handling

- **Placeholder Creation**: Automatically creates colored placeholder images if originals are missing
- **Proper Scaling**: Images are scaled to fit within slide bounds while maintaining aspect ratio
- **Multiple Images**: Supports multiple images per slide with grid layout

## API Reference

### PowerPointAPI Class

#### CreatePresentationFromJsonFile
```csharp
public static async Task<string> CreatePresentationFromJsonFile(
    string jsonFilePath,
    string outputPath,
    string presentationTitle = "JSON Generated Presentation",
    string author = "AI Assistant",
    string? imageBasePath = null)
```

#### CreatePresentationFromJsonString
```csharp
public static async Task<string> CreatePresentationFromJsonString(
    string jsonContent,
    string outputPath,
    string presentationTitle = "JSON Generated Presentation",
    string author = "AI Assistant",
    string? imageBasePath = null)
```

### JsonSlideParser Class

#### ParseFromFile
```csharp
public static PresentationContent ParseFromFile(
    string jsonFilePath,
    string presentationTitle = "JSON Generated Presentation",
    string author = "AI Assistant",
    string? imageBasePath = null)
```

#### ParseFromString
```csharp
public static PresentationContent ParseFromString(
    string jsonContent,
    string presentationTitle = "JSON Generated Presentation", 
    string author = "AI Assistant",
    string? imageBasePath = null)
```

## Dependencies

- .NET 8.0
- DocumentFormat.OpenXml (Open-XML-SDK)
- System.Drawing.Common (for image processing)
- System.Text.Json (for JSON parsing)

## Error Handling

The application includes comprehensive error handling for:
- Missing JSON files
- Invalid JSON format
- Missing image files (creates placeholders)
- OpenXML generation errors

## Output

Generated PowerPoint files are fully compatible with:
- Microsoft PowerPoint 2013 and later
- PowerPoint Online
- LibreOffice Impress
- Google Slides (with import)

## Example Output

Running the program with the default `slides_content.json` creates:
- `Can_Opener_Presentation.pptx` - Main presentation file
- `Images/` folder with placeholder images (if originals are missing)

## License

This project is provided as-is for educational and commercial use.
