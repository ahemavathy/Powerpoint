# PowerPoint Generator Web Service - Architecture Design

## System Overview

The PowerPoint Generator Web Service is a RESTful API built with ASP.NET Core that generates PowerPoint presentations from JSON content and manages image assets. The system follows a layered architecture pattern with clear separation of concerns.

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Client Layer                             │
├─────────────────────────────────────────────────────────────────┤
│  Web Browsers  │  Mobile Apps  │  Desktop Apps  │  Other APIs   │
│  (Swagger UI)  │              │               │               │
└─────────────────┬───────────────┬───────────────┬───────────────┘
                  │               │               │
                  └───────────────┼───────────────┘
                                  │
                         HTTP/HTTPS REST API
                                  │
┌─────────────────────────────────┼─────────────────────────────────┐
│                    API Gateway Layer                             │
├─────────────────────────────────┼─────────────────────────────────┤
│           ASP.NET Core Web API                                   │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │               Presentation Controller                       │ │
│  │  • Image Upload/Management                                 │ │
│  │  • Presentation Creation                                   │ │
│  │  • File Download/Management                               │ │
│  │  • Health Check                                           │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                              │                                   │
├──────────────────────────────┼───────────────────────────────────┤
│                    Middleware Layer                              │
├──────────────────────────────┼───────────────────────────────────┤
│  • CORS Handling            │  • Error Handling                 │
│  • Authentication (future)   │  • Logging & Monitoring          │
│  • Rate Limiting (future)    │  • Request/Response Validation   │
└──────────────────────────────┼───────────────────────────────────┘
                               │
┌──────────────────────────────┼───────────────────────────────────┐
│                    Business Logic Layer                          │
├──────────────────────────────┼───────────────────────────────────┤
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                    Services                                 │ │
│  │  ┌─────────────────────┬─────────────────────────────────┐  │ │
│  │  │ PowerPointGenerator │      JsonSlideParser           │  │ │
│  │  │ Service             │      Service                    │  │ │
│  │  │                     │                                 │  │ │
│  │  │ • Presentation      │ • JSON Content Parsing         │  │ │
│  │  │   Creation          │ • Slide Content Extraction     │  │ │
│  │  │ • OpenXML           │ • Image Reference Resolution   │  │ │
│  │  │   Manipulation      │ • Data Validation              │  │ │
│  │  └─────────────────────┴─────────────────────────────────┘  │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                              │                                   │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                   Utilities                                 │ │
│  │  ┌─────────────────┬─────────────────┬─────────────────────┐ │ │
│  │  │   SlideHelper   │   ImageHelper   │    ThemeHelper      │ │ │
│  │  │                 │                 │                     │ │ │
│  │  │ • Slide Layout  │ • Image         │ • Theme Creation    │ │ │
│  │  │ • Text Formatting│   Processing   │ • Style Management  │ │ │
│  │  │ • Shape Creation│ • Placeholder   │ • Color Schemes     │ │ │
│  │  └─────────────────┴─────────────────┴─────────────────────┘ │ │
│  └─────────────────────────────────────────────────────────────┘ │
└──────────────────────────────┼───────────────────────────────────┘
                               │
┌──────────────────────────────┼───────────────────────────────────┐
│                     Data Access Layer                            │
├──────────────────────────────┼───────────────────────────────────┤
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                   File System                               │ │
│  │  ┌─────────────────────┬─────────────────────────────────┐  │ │
│  │  │  Generated          │         Images                  │  │ │
│  │  │  Presentations/     │         Directory               │  │ │
│  │  │                     │                                 │  │ │
│  │  │ • .pptx files       │ • Uploaded images               │  │ │
│  │  │ • Metadata          │ • Image processing              │  │ │
│  │  │ • Cleanup           │ • Format validation             │  │ │
│  │  └─────────────────────┴─────────────────────────────────┘  │ │
│  └─────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
```

## Component Architecture

### 1. API Layer Components

#### PresentationController
- **Responsibility**: HTTP request handling and response formatting
- **Key Methods**:
  - `POST /api/presentation/create-from-json` - Create presentations
  - `POST /api/presentation/upload-image` - Single image upload
  - `POST /api/presentation/upload-images` - Multiple image upload
  - `GET /api/presentation/images` - List uploaded images
  - `GET /api/presentation/image/{fileName}` - Retrieve specific image
  - `GET /api/presentation/list` - List generated presentations
  - `GET /api/presentation/download/{fileName}` - Download presentation
  - `DELETE /api/presentation/delete/{fileName}` - Delete presentation
  - `DELETE /api/presentation/image/{fileName}` - Delete image
  - `GET /api/presentation/health` - Health check

### 2. Service Layer Components

#### PowerPointGeneratorService
```csharp
Responsibilities:
├── Presentation Creation
│   ├── OpenXML document initialization
│   ├── Slide generation and layout
│   └── Content population
├── Resource Management
│   ├── Image embedding
│   ├── Theme application
│   └── File cleanup
└── Error Handling
    ├── OpenXML exceptions
    ├── File I/O errors
    └── Content validation
```

#### JsonSlideParser
```csharp
Responsibilities:
├── JSON Parsing
│   ├── Content validation
│   ├── Schema verification
│   └── Error handling
├── Data Transformation
│   ├── JSON to domain models
│   ├── Image path resolution
│   └── Content sanitization
└── Configuration
    ├── Default values
    ├── Presentation metadata
    └── Slide ordering
```

### 3. Utility Layer Components

#### SlideHelper
- Slide layout management
- Text formatting and positioning
- Shape and placeholder creation
- OpenXML element generation

#### ImageHelper
- Image processing and validation
- Placeholder image generation
- Format conversion support
- Dimension extraction

#### ThemeHelper
- Theme creation and management
- Color scheme application
- Font and style configuration
- Corporate branding support

### 4. Model Layer

#### Domain Models
```csharp
PresentationContent
├── Metadata (title, author, created date)
├── Slides[] (collection of slide content)
└── Configuration (theme, layout preferences)

SlideContent
├── Title (slide heading)
├── Description (main content)
├── ImagePath (reference to image file)
└── Layout (positioning and formatting)

ImageContent
├── FilePath (local file system path)
├── PlaceholderText (fallback content)
└── Dimensions (width, height)
```

#### API Models
```csharp
CreatePresentationRequest
├── JsonContent (slide data as JSON string)
├── PresentationName (optional custom name)
├── PresentationTitle (display title)
└── Author (presentation author)

ImageUploadResponse
├── Success (operation result)
├── FileName (uploaded file name)
├── FileSize (file size in bytes)
├── ImageUrl (access URL)
└── ErrorMessage (failure details)
```

## Data Flow Architecture

### 1. Presentation Creation Flow
```
Client Request → Controller → JSON Parser → Domain Models → 
PowerPoint Generator → OpenXML Processing → File System → Response
```

### 2. Image Upload Flow
```
Client Upload → Controller → Validation → File System Storage → 
Metadata Extraction → Response Generation
```

### 3. File Management Flow
```
Client Request → Controller → File System Operations → 
Response/Stream Generation
```

## Security Architecture

### Current Implementation
- **CORS**: Enabled for cross-origin requests
- **Input Validation**: File type and size restrictions
- **Error Handling**: Sanitized error responses
- **File Access**: Restricted to designated directories

### Future Security Enhancements
```
Authentication Layer
├── JWT Token Validation
├── API Key Management
└── Role-Based Access Control

Authorization Layer
├── Resource-Level Permissions
├── Rate Limiting
└── Audit Logging

Data Protection
├── Input Sanitization
├── SQL Injection Prevention
└── File Upload Security
```

## Scalability Architecture

### Current Design
- **Stateless**: No server-side session management
- **File-Based**: Simple file system storage
- **Single Instance**: Designed for single server deployment

### Scalability Considerations
```
Horizontal Scaling
├── Load Balancer Integration
├── Shared File Storage (NFS/S3)
└── Database Migration

Performance Optimization
├── Caching Layer (Redis)
├── Background Processing (Queues)
└── CDN Integration

Monitoring & Observability
├── Application Performance Monitoring
├── Health Check Endpoints
└── Structured Logging
```

## Deployment Architecture

### Current Deployment
```
Single Server Deployment
├── ASP.NET Core Application
├── Local File System
└── In-Process Request Handling
```

### Production Deployment Options

#### Option 1: Container Deployment
```
Docker Container
├── ASP.NET Core Runtime
├── Application Code
├── Volume Mounts
│   ├── /app/Images (persistent storage)
│   └── /app/GeneratedPresentations
└── Health Check Configuration
```

#### Option 2: Cloud Deployment
```
Azure App Service
├── Application Hosting
├── Azure Storage Account
│   ├── Blob Storage (images)
│   └── File Share (presentations)
├── Application Insights (monitoring)
└── Azure CDN (static content)
```

#### Option 3: Microservices Architecture
```
API Gateway
├── Presentation Service
├── Image Management Service
├── File Storage Service
└── Notification Service
```

## Technology Stack

### Core Technologies
- **Framework**: ASP.NET Core 8.0
- **Language**: C# 12
- **Documentation**: OpenAPI/Swagger
- **File Processing**: DocumentFormat.OpenXml
- **Image Processing**: System.Drawing.Common

### Dependencies
```
Production Dependencies
├── DocumentFormat.OpenXml (PowerPoint generation)
├── System.Drawing.Common (image processing)
├── Microsoft.AspNetCore.OpenApi (API documentation)
└── Swashbuckle.AspNetCore (Swagger UI)

Development Dependencies
├── Microsoft.Extensions.Logging (logging)
├── System.Text.Json (JSON processing)
└── Microsoft.AspNetCore.Mvc (MVC framework)
```

## Performance Characteristics

### Current Performance Profile
- **Memory Usage**: ~50-100MB base + 10-20MB per concurrent request
- **CPU Usage**: High during PowerPoint generation, low at rest
- **Disk I/O**: Moderate (file uploads/downloads)
- **Network**: Dependent on file sizes and concurrent users

### Performance Optimization Strategies
1. **Async Processing**: All I/O operations are asynchronous
2. **Memory Management**: Using statements for proper disposal
3. **File Streaming**: Direct file streaming for downloads
4. **Error Caching**: Quick failure for invalid requests

## Monitoring and Observability

### Health Monitoring
- **Health Check Endpoint**: `/api/presentation/health`
- **Service Status**: Application health verification
- **Dependency Checks**: File system access validation

### Logging Strategy
```
Log Levels
├── Information (successful operations)
├── Warning (non-critical issues)
├── Error (operation failures)
└── Critical (system failures)

Log Categories
├── API Requests/Responses
├── File Operations
├── Business Logic Events
└── System Performance
```

## Future Architecture Enhancements

### Phase 1: Enhanced Features
- Template system for presentation designs
- Batch processing for multiple presentations
- Webhook notifications for completion events

### Phase 2: Enterprise Features
- Multi-tenant architecture
- Advanced authentication/authorization
- API versioning and backward compatibility

### Phase 3: AI Integration
- Content suggestion engine
- Image recommendation system
- Layout optimization algorithms

This architecture provides a solid foundation for the current PowerPoint Generator web service while allowing for future scalability and feature enhancements.
