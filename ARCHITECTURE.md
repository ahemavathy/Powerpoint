# PowerPoint Generator - Architecture & Technical Specifications

## Overview

The PowerPoint Generator is a comprehensive C# .NET 8.0 solution that provides both a console application and RESTful Web API for generating PowerPoint presentations from JSON content. The system uses the DocumentFormat.OpenXml SDK and follows a layered architecture pattern with clear separation of concerns.

## System Requirements

### Runtime Environment
- **.NET Version**: 8.0
- **Operating System**: Windows (primary), Linux/macOS (limited due to System.Drawing.Common)
- **Memory**: Minimum 512MB, Recommended 1GB+
- **Storage**: 100MB application + storage for generated files
- **CPU**: Multi-core recommended for concurrent processing

### Dependencies

**Console Application** (`PowerPointGenerator.csproj`):
```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
<PackageReference Include="System.Drawing.Common" Version="9.0.7" />
<PackageReference Include="System.Text.Json" Version="8.0.0" />
```

**Web API** (`WebAPI/PowerPointGenerator.WebAPI.csproj`):
```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.0.1" />
<PackageReference Include="System.Drawing.Common" Version="8.0.0" />
<PackageReference Include="Microsoft.AspNetCore.OpenApi" Version="8.0.0" />
<PackageReference Include="Swashbuckle.AspNetCore" Version="6.4.0" />
```

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Client Applications                      │
├─────────────────────┬───────────────────────────────────────┤
│    Console App      │           Web Applications            │
│                     │    (Browser, Mobile, API Clients)    │
└─────────────────────┴───────────────────────────────────────┘
           │                              │
           │                              │
           ▼                              ▼
┌─────────────────────┐         ┌─────────────────────────────┐
│   Console Interface │         │      Web API Layer         │
│                     │         │   (ASP.NET Core 8.0)       │
│   Program.cs        │         │   Controllers + Swagger    │
│   PowerPointAPI.cs  │         └─────────────────────────────┘
└─────────────────────┘                        │
           │                                   │
           └───────────────┬───────────────────┘
                          │
                          ▼
        ┌─────────────────────────────────────────────────────┐
        │              Business Logic Layer                   │
        │                                                     │
        │  ┌─────────────────┐  ┌─────────────────────────┐   │
        │  │    Services     │  │       Utilities         │   │
        │  │                 │  │                         │   │
        │  │ • PowerPoint    │  │ • SlideHelper           │   │
        │  │   Generator     │  │ • ImageHelper           │   │
        │  │ • JSON Parser   │  │ • ThemeHelper           │   │
        │  │ • Slide Content │  │                         │   │
        │  │   Parser        │  │                         │   │
        │  └─────────────────┘  └─────────────────────────┘   │
        └─────────────────────────────────────────────────────┘
                          │
                          ▼
        ┌─────────────────────────────────────────────────────┐
        │               Data Access Layer                     │
        │                                                     │
        │  ┌─────────────────┐  ┌─────────────────────────┐   │
        │  │  File System    │  │    OpenXML SDK          │   │
        │  │                 │  │                         │   │
        │  │ • Images/       │  │ • Document Creation     │   │
        │  │ • Generated     │  │ • Slide Management      │   │
        │  │   Presentations │  │ • Image Embedding       │   │
        │  │ • JSON Input    │  │ • Theme Application     │   │
        │  └─────────────────┘  └─────────────────────────┘   │
        └─────────────────────────────────────────────────────┘
```

## Component Architecture

### Core Components

#### 1. **Models Layer** (`Models/`)
- **PresentationModels.cs**: Core domain models (PresentationContent, SlideContent, ImageContent)
- **JsonSlideModels.cs**: JSON-specific input models (JsonSlideContent, JsonSlide)
- **WebApiModels.cs**: API request/response models (CreatePresentationRequest, PresentationResponse)

#### 2. **Services Layer** (`Services/`)
- **PowerPointGeneratorService.cs**: Core presentation generation logic using OpenXML, including template-based generation with placeholder replacement and content removal
- **JsonSlideParser.cs**: Parses JSON content into domain models with layout type support
- **SlideContentParser.cs**: Legacy text-based content parsing

#### 3. **Utilities Layer** (`Utilities/`)
- **SlideHelper.cs**: Slide creation and formatting utilities
- **ImageHelper.cs**: Image processing with aspect ratio preservation
- **ThemeHelper.cs**: PowerPoint theme and styling management

#### 4. **API Layer** (`Controllers/` & `WebAPI/`)
- **PresentationController.cs**: REST API endpoints
- **Program.cs**: Web API configuration and startup
- **PowerPointAPI.cs**: Simplified API wrapper for console usage

## Data Flow Architecture

### Console Application Flow
```
JSON File → JsonSlideParser → PresentationContent → PowerPointGeneratorService → .pptx File
```

### Template-Based Generation Flow
```
Template.pptx + JSON Content → PowerPointGeneratorService → Modified Template → .pptx File
```

### Web API Flow
```
HTTP Request → Controller → JsonSlideParser → PowerPointGeneratorService → File Storage → HTTP Response
```

### Template Processing Flow
```
1. Copy template to output location
2. Replace text placeholders ({title}, {description}) with JSON content
3. Replace image placeholders with uploaded images
4. Remove excess slides if JSON has fewer slides than template
5. Remove image parts from slides if JSON content has no image
```

## Technical Specifications

### API Endpoints

**Base URL**: `http://localhost:5000/api/presentation`

#### Core Endpoints
- `POST /create-from-json` - Create presentation from JSON content
- `POST /create-from-template` - Create presentation from PowerPoint template with placeholder replacement
- `GET /download/{fileName}` - Download generated presentation
- `GET /list` - List all presentations
- `DELETE /delete/{fileName}` - Delete presentation

#### Image Management
- `POST /upload-image` - Upload single image
- `POST /upload-images` - Upload multiple images
- `GET /images` - List uploaded images
- `GET /image/{fileName}` - Get image file
- `DELETE /image/{fileName}` - Delete image

#### Template Management
- `POST /upload-template` - Upload PowerPoint template file (.pptx)
- `GET /templates` - List uploaded templates
- `GET /template/{fileName}` - Get template file
- `DELETE /template/{fileName}` - Delete template

#### Utility
- `GET /health` - Health check

### Data Models

#### JSON Input Format
```json
{
  "slides": [
    {
      "title": "Slide Title",
      "description": "Slide description content",
      "suggested_image": "image-filename.png",
      "layout": "SingleImage"
    }
  ]
}
```

#### API Request Format
```json
{
  "jsonContent": "{\"slides\": [...]}",
  "presentationName": "MyPresentation",
  "presentationTitle": "My Presentation Title",
  "author": "Author Name"
}
```

#### Template-Based Generation Request Format
```json
{
  "jsonContent": "{\"slides\": [...]}",
  "templatePath": "Templates/MyTemplate.pptx",
  "presentationName": "Generated_Presentation"
}
```

### File Management

#### Directory Structure
```
Application Root/
├── Images/                          # Console app images
├── Templates/                       # PowerPoint template files (.pptx)
├── slides_content.json              # Default input file
├── WebAPI/
│   ├── Images/                      # API uploaded images
│   ├── Templates/                   # API uploaded templates
│   ├── GeneratedPresentations/      # API generated files
│   └── Program.cs
└── Generated files (console output)
```

#### File Validation
- **Supported Image Formats**: JPG, JPEG, PNG, GIF, BMP, WEBP
- **Template Format**: PowerPoint (.pptx) files only
- **Maximum File Size**: 10MB per image, 50MB per template
- **Presentation Format**: Office Open XML (.pptx)

### Performance Specifications

#### Response Time Targets
- **Health Check**: < 50ms
- **Image Upload**: < 2 seconds (10MB file)
- **Presentation Creation**: < 5 seconds (10 slides)
- **Console Generation**: < 3 seconds (typical)

#### Resource Limits
- **Slides per Presentation**: 100 maximum
- **Images per Slide**: 1 (current implementation)
- **Processing Timeout**: 30 seconds
- **Memory Usage**: 50-200MB per operation

### Image Processing Features

#### Aspect Ratio Preservation
The system maintains original image proportions using advanced scaling algorithms:

```csharp
// From ImageHelper.cs
public static (long width, long height) CalculateFitDimensions(
    int originalWidth, int originalHeight, 
    long maxWidth, long maxHeight)
```

#### Template-Based Generation Features

##### Placeholder Replacement
The system automatically replaces text placeholders in template slides:
- `{title}` - Replaced with slide title from JSON content
- `{description}` - Replaced with slide description from JSON content

##### Smart Content Management
- **Excess Slide Removal**: If JSON content has fewer slides than the template, extra slides are automatically removed
- **Image Part Removal**: If a slide in JSON content has no image, image parts are removed from the corresponding template slide
- **Layout Preservation**: Original template formatting, themes, and layout structure are maintained

##### Template Structure Requirements
Templates should follow these conventions:
- Use `{title}` and `{description}` as text placeholders
- Include image placeholders for slides that will contain images
- Maintain consistent slide structure for best results

#### Layout Types
- **Single Large Image**: Maximum space utilization below text
- **Product Showcase**: Side-by-side layout with text on left, image on right
- **Image Grid**: 2x2 grid for multiple images (future enhancement)
- **Image with Caption**: Detailed image descriptions (future enhancement)
- **Two-Image Comparison**: Side-by-side layout (future enhancement)
- **Template-Based**: Custom layouts using PowerPoint template files with placeholder replacement

## Security Architecture

### Current Security Measures
- **Input Validation**: File type and size validation
- **Path Security**: Prevention of directory traversal
- **CORS Policy**: Configurable cross-origin requests
- **Error Handling**: Sanitized error responses

### Error Handling
```csharp
HTTP Status Codes:
├── 200 OK - Successful operation
├── 400 Bad Request - Invalid input
├── 404 Not Found - Resource not found
├── 413 Payload Too Large - File size exceeded
└── 500 Internal Server Error - System error
```

## Deployment Architecture

### Console Application
```bash
# Run directly
dotnet run --project PowerPointGenerator.csproj

# With arguments
dotnet run --project PowerPointGenerator.csproj slides.json "My Presentation"
```

### Web API
```bash
# Development
dotnet run --project WebAPI/PowerPointGenerator.WebAPI.csproj

# Production
dotnet publish WebAPI/ -c Release -o ./publish
cd publish && dotnet PowerPointGenerator.WebAPI.dll
```

### Docker Deployment
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY ./publish .
EXPOSE 80
VOLUME ["/app/Images", "/app/GeneratedPresentations"]
ENTRYPOINT ["dotnet", "PowerPointGenerator.WebAPI.dll"]
```

## Configuration

### Application Settings
```json
{
  "PowerPointGenerator": {
    "MaxFileSize": 10485760,
    "MaxTemplateSize": 52428800,
    "AllowedImageTypes": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"],
    "AllowedTemplateTypes": [".pptx"],
    "DefaultTheme": "Office",
    "ProcessingTimeout": 30000,
    "TemplateDirectory": "Templates",
    "PlaceholderPatterns": {
      "Title": "{title}",
      "Description": "{description}",
      "Image": "placeholder-image"
    }
  }
}
```

### Environment Variables
```bash
ASPNETCORE_ENVIRONMENT=Development|Production
ASPNETCORE_URLS=http://localhost:5000
POWERPOINT_STORAGE_PATH=/app/storage
```

## Future Enhancements

### Planned Features
- **Authentication**: JWT-based API security
- **Multi-tenancy**: User isolation and resource management
- **Advanced Layouts**: More slide templates and designs
- **Enhanced Template Features**: Dynamic placeholder discovery, template validation, layout optimization
- **Real-time Processing**: WebSocket-based progress updates
- **Cloud Storage**: Azure Blob Storage integration
- **Batch Processing**: Multiple presentation generation
- **AI Integration**: Content suggestions and optimization
- **Template Gallery**: Pre-built template library with categories

### Scalability Improvements
- **Horizontal Scaling**: Load balancer support
- **Caching**: Redis-based result caching
- **Queue Processing**: Background job processing
- **Microservices**: Service decomposition for enterprise scale

## Monitoring and Maintenance

### Logging
- **Console App**: Standard output logging
- **Web API**: ASP.NET Core structured logging
- **Error Tracking**: Comprehensive exception handling

### Health Monitoring
- **API Health Endpoint**: `/api/presentation/health`
- **File System Monitoring**: Storage space and permissions
- **Performance Metrics**: Response times and throughput

---

This comprehensive architecture and technical specification document provides the complete technical foundation for both the console application and web API components of the PowerPoint Generator system.
