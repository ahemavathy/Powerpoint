# PowerPoint Generator Web Service - Technical Specifications

## System Requirements

### Runtime Environment
- **.NET Version**: 8.0 or higher
- **Operating System**: Windows, Linux, macOS
- **Memory**: Minimum 512MB, Recommended 2GB+
- **Storage**: 100MB application + storage for generated files
- **CPU**: Multi-core recommended for concurrent processing

### Dependencies
```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
<PackageReference Include="System.Drawing.Common" Version="9.0.7" />
<PackageReference Include="Microsoft.AspNetCore.OpenApi" Version="8.0.0" />
<PackageReference Include="Swashbuckle.AspNetCore" Version="6.4.0" />
```

## API Specifications

### Base Configuration
- **Base URL**: `http://localhost:5000/api/presentation`
- **Content-Type**: `application/json` for JSON endpoints, `multipart/form-data` for file uploads
- **Response Format**: JSON
- **HTTP Methods**: GET, POST, DELETE

### Endpoint Specifications

#### 1. Health Check
```http
GET /api/presentation/health
```
**Response**:
```json
{
  "status": "healthy",
  "timestamp": "2025-01-22T14:30:52Z",
  "version": "1.0.0",
  "service": "PowerPoint Generator API"
}
```

#### 2. Create Presentation
```http
POST /api/presentation/create-from-json
Content-Type: application/json
```
**Request Body**:
```json
{
  "jsonContent": "{\\"slides\\": [{\\"title\\": \\"Slide Title\\", \\"description\\": \\"Slide content\\", \\"suggested_image\\": \\"image.png\\"}]}",
  "presentationName": "MyPresentation",
  "presentationTitle": "My Presentation Title",
  "author": "John Doe"
}
```
**Response**:
```json
{
  "success": true,
  "fileName": "MyPresentation_abc123.pptx",
  "filePath": "/full/path/to/file.pptx",
  "presentationName": "MyPresentation",
  "createdAt": "2025-01-22T14:30:52Z",
  "fileSize": 1234567,
  "slideCount": 4,
  "downloadUrl": "/api/presentation/download/MyPresentation_abc123.pptx"
}
```

#### 3. Upload Image
```http
POST /api/presentation/upload-image
Content-Type: multipart/form-data
```
**Form Data**:
- `file`: Image file (JPG, JPEG, PNG, GIF, BMP, WEBP)

**Response**:
```json
{
  "success": true,
  "fileName": "image.png",
  "filePath": "/full/path/to/image.png",
  "uploadedAt": "2025-01-22T14:30:52Z",
  "fileSize": 1234567,
  "imageUrl": "/api/presentation/image/image.png"
}
```

#### 4. List Images
```http
GET /api/presentation/images
```
**Response**:
```json
[
  {
    "fileName": "image.png",
    "uploadedAt": "2025-01-22T14:30:52Z",
    "fileSize": 1234567,
    "imageUrl": "/api/presentation/image/image.png",
    "dimensions": "1920 x 1080"
  }
]
```

#### 5. Download Files
```http
GET /api/presentation/download/{fileName}
GET /api/presentation/image/{fileName}
```
**Response**: Binary file stream with appropriate content-type headers

## Data Models

### Domain Models
```csharp
public class PresentationContent
{
    public string Title { get; set; }
    public string Author { get; set; }
    public DateTime CreatedDate { get; set; }
    public List<SlideContent> Slides { get; set; }
    public string ThemeName { get; set; }
}

public class SlideContent
{
    public string Title { get; set; }
    public string Description { get; set; }
    public ImageContent? Image { get; set; }
    public SlideLayout Layout { get; set; }
}

public class ImageContent
{
    public string FilePath { get; set; }
    public string PlaceholderText { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
}
```

### JSON Input Format
```json
{
  "slides": [
    {
      "title": "Slide Title",
      "description": "Slide description content",
      "suggested_image": "image-filename.png"
    }
  ]
}
```

## File Management Specifications

### Directory Structure
```
Application Root/
├── Images/                     # Uploaded image files
│   ├── image1.png
│   ├── image2.jpg
│   └── ...
├── GeneratedPresentations/     # Generated PowerPoint files
│   ├── presentation1.pptx
│   ├── presentation2.pptx
│   └── ...
└── WebAPI/                     # Web API application files
    ├── Program.cs
    ├── Properties/
    └── ...
```

### File Naming Conventions
- **Images**: Original filename (exact as uploaded)
- **Presentations**: `{PresentationName}_{UniqueID}.pptx`
- **Unique ID**: GUID without hyphens (32 characters)

### File Validation Rules
```csharp
Image Files:
├── Allowed Extensions: .jpg, .jpeg, .png, .gif, .bmp, .webp
├── Maximum Size: 10MB
├── Validation: File header verification
└── Processing: Dimension extraction

PowerPoint Files:
├── Format: Office Open XML (.pptx)
├── Compatibility: PowerPoint 2016+
├── Encoding: UTF-8
└── Structure: Valid OpenXML document
```

## Performance Specifications

### Response Time Targets
- **Health Check**: < 50ms
- **Image Upload**: < 2 seconds (10MB file)
- **Presentation Creation**: < 5 seconds (10 slides)
- **File Download**: Streaming (no fixed time)
- **List Operations**: < 500ms

### Throughput Specifications
- **Concurrent Users**: 10-50 (single instance)
- **Requests per Second**: 10-100 (varies by operation)
- **File Processing**: 1-5 presentations simultaneously
- **Memory Usage**: 50-200MB per active request

### Resource Limits
```yaml
File Limits:
  Image File Size: 10MB maximum
  Presentation Size: 100MB maximum
  Concurrent Uploads: 5 per user
  Total Storage: Configurable (default: unlimited)

Processing Limits:
  Slides per Presentation: 100 maximum
  Images per Slide: 1
  Text Length: 10,000 characters per slide
  Processing Timeout: 30 seconds
```

## Error Handling Specifications

### HTTP Status Codes
```
200 OK - Successful operation
400 Bad Request - Invalid input data
404 Not Found - Resource not found
413 Payload Too Large - File size exceeded
422 Unprocessable Entity - Validation failed
500 Internal Server Error - Server error
```

### Error Response Format
```json
{
  "error": "Error category",
  "details": "Detailed error message",
  "timestamp": "2025-01-22T14:30:52Z",
  "traceId": "unique-trace-identifier"
}
```

### Error Categories
```csharp
Validation Errors:
├── Invalid JSON format
├── Missing required fields
├── File type not supported
└── File size exceeded

Processing Errors:
├── PowerPoint generation failed
├── Image processing failed
├── File system errors
└── OpenXML format errors

System Errors:
├── Out of memory
├── Disk space insufficient
├── Service unavailable
└── Timeout exceeded
```

## Security Specifications

### Current Security Measures
```yaml
CORS Policy:
  Allow Origins: "*" (configurable)
  Allow Methods: GET, POST, DELETE
  Allow Headers: "*"
  Credentials: false

File Upload Security:
  Type Validation: File extension and MIME type
  Size Limits: 10MB per file
  Path Traversal: Prevented
  Virus Scanning: Not implemented (future)

Input Validation:
  JSON Schema: Strict validation
  SQL Injection: Not applicable (no database)
  XSS Protection: Output encoding
  CSRF Protection: Not implemented (stateless API)
```

### Future Security Enhancements
```yaml
Authentication:
  Method: JWT Bearer tokens
  Provider: Configurable (Azure AD, Auth0, etc.)
  Token Expiry: 1 hour default
  Refresh Tokens: Supported

Authorization:
  Model: Role-based access control (RBAC)
  Roles: Admin, User, ReadOnly
  Permissions: Resource-level access control
  API Keys: For service-to-service authentication

Data Protection:
  Encryption: TLS 1.2+ for transport
  At Rest: File system encryption
  Logging: Sensitive data masking
  Audit Trail: All operations logged
```

## Configuration Specifications

### Application Settings
```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*",
  "PowerPointGenerator": {
    "MaxFileSize": 10485760,
    "AllowedImageTypes": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"],
    "DefaultTheme": "Office",
    "ProcessingTimeout": 30000,
    "CleanupInterval": 3600000
  }
}
```

### Environment Variables
```bash
ASPNETCORE_ENVIRONMENT=Development|Production
ASPNETCORE_URLS=http://localhost:5000
POWERPOINT_STORAGE_PATH=/app/storage
POWERPOINT_MAX_FILE_SIZE=10485760
POWERPOINT_CLEANUP_ENABLED=true
```

## Monitoring and Logging Specifications

### Logging Configuration
```csharp
Log Levels:
├── Trace: Detailed execution flow
├── Debug: Development debugging info
├── Information: General application flow
├── Warning: Unexpected but handled events
├── Error: Application errors and exceptions
└── Critical: System failures requiring immediate attention

Log Categories:
├── PowerPointGenerator.Controllers.*
├── PowerPointGenerator.Services.*
├── PowerPointGenerator.Utilities.*
└── Microsoft.AspNetCore.*
```

### Metrics Collection
```yaml
Performance Metrics:
  - Request duration
  - Memory usage
  - CPU utilization
  - File system I/O
  - Error rates
  - Throughput (requests/second)

Business Metrics:
  - Presentations created
  - Images uploaded
  - Files downloaded
  - Storage usage
  - User activity
  - Feature usage
```

## Deployment Specifications

### Production Deployment Checklist
```yaml
Environment Setup:
  ☐ .NET 8.0 Runtime installed
  ☐ Required directories created
  ☐ File permissions configured
  ☐ Storage space allocated
  ☐ Network ports opened (80, 443)

Security Configuration:
  ☐ HTTPS certificate installed
  ☐ CORS policy configured
  ☐ File upload limits set
  ☐ Error handling configured
  ☐ Logging configured

Performance Tuning:
  ☐ Memory limits configured
  ☐ Request timeout settings
  ☐ File cleanup scheduled
  ☐ Monitoring enabled
  ☐ Health checks configured
```

### Container Specifications
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

# Application files
COPY --from=publish /app/publish .

# Storage volumes
VOLUME ["/app/Images", "/app/GeneratedPresentations"]

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost/api/presentation/health || exit 1

ENTRYPOINT ["dotnet", "PowerPointGenerator.WebAPI.dll"]
```

This technical specification provides comprehensive details for implementing, deploying, and maintaining the PowerPoint Generator web service.
