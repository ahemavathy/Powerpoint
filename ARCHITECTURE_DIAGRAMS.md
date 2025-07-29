# PowerPoint Generator - System Architecture Diagram

## Component Relationship Diagram

```mermaid
graph TB
    %% Client Layer
    subgraph "Client Layer"
        WEB[Web Browser<br/>Swagger UI]
        MOBILE[Mobile Apps]
        DESKTOP[Desktop Apps]
        API_CLIENT[Other APIs<br/>Integrations]
    end

    %% API Gateway
    subgraph "ASP.NET Core Web API"
        CONTROLLER[PresentationController<br/>• Image Management<br/>• Presentation Creation<br/>• File Operations<br/>• Health Checks]
        
        subgraph "Middleware"
            CORS[CORS Handler]
            AUTH[Authentication<br/>Future]
            LOGGING[Logging & Monitoring]
            ERROR[Error Handling]
        end
    end

    %% Business Logic Layer
    subgraph "Business Logic Layer"
        subgraph "Services"
            PPT_SERVICE[PowerPointGeneratorService<br/>• OpenXML Processing<br/>• Presentation Creation<br/>• Resource Management]
            JSON_PARSER[JsonSlideParser<br/>• JSON Validation<br/>• Content Parsing<br/>• Data Transformation]
        end
        
        subgraph "Utilities"
            SLIDE_HELPER[SlideHelper<br/>• Layout Management<br/>• Text Formatting<br/>• Shape Creation]
            IMAGE_HELPER[ImageHelper<br/>• Image Processing<br/>• Placeholder Generation<br/>• Format Validation]
            THEME_HELPER[ThemeHelper<br/>• Theme Creation<br/>• Style Management<br/>• Color Schemes]
        end
    end

    %% Data Layer
    subgraph "Data Access Layer"
        subgraph "File System"
            IMAGES_DIR[Images Directory<br/>• Uploaded Images<br/>• Format Validation<br/>• Metadata Extraction]
            PPT_DIR[Presentations Directory<br/>• Generated .pptx Files<br/>• File Metadata<br/>• Cleanup Management]
        end
    end

    %% Model Layer
    subgraph "Models"
        DOMAIN_MODELS[Domain Models<br/>• PresentationContent<br/>• SlideContent<br/>• ImageContent]
        API_MODELS[API Models<br/>• CreatePresentationRequest<br/>• ImageUploadResponse<br/>• PresentationResponse]
        JSON_MODELS[JSON Models<br/>• JsonSlideContent<br/>• JsonSlide]
    end

    %% Connections
    WEB --> CONTROLLER
    MOBILE --> CONTROLLER
    DESKTOP --> CONTROLLER
    API_CLIENT --> CONTROLLER

    CONTROLLER --> CORS
    CONTROLLER --> AUTH
    CONTROLLER --> LOGGING
    CONTROLLER --> ERROR

    CONTROLLER --> PPT_SERVICE
    CONTROLLER --> JSON_PARSER

    PPT_SERVICE --> SLIDE_HELPER
    PPT_SERVICE --> IMAGE_HELPER
    PPT_SERVICE --> THEME_HELPER

    JSON_PARSER --> DOMAIN_MODELS
    PPT_SERVICE --> DOMAIN_MODELS

    CONTROLLER --> API_MODELS
    API_MODELS --> JSON_MODELS

    PPT_SERVICE --> PPT_DIR
    CONTROLLER --> IMAGES_DIR
    IMAGE_HELPER --> IMAGES_DIR

    classDef clientStyle fill:#e1f5fe,stroke:#01579b,stroke-width:2px
    classDef apiStyle fill:#f3e5f5,stroke:#4a148c,stroke-width:2px
    classDef serviceStyle fill:#e8f5e8,stroke:#1b5e20,stroke-width:2px
    classDef dataStyle fill:#fff3e0,stroke:#e65100,stroke-width:2px
    classDef modelStyle fill:#fce4ec,stroke:#880e4f,stroke-width:2px

    class WEB,MOBILE,DESKTOP,API_CLIENT clientStyle
    class CONTROLLER,CORS,AUTH,LOGGING,ERROR apiStyle
    class PPT_SERVICE,JSON_PARSER,SLIDE_HELPER,IMAGE_HELPER,THEME_HELPER serviceStyle
    class IMAGES_DIR,PPT_DIR dataStyle
    class DOMAIN_MODELS,API_MODELS,JSON_MODELS modelStyle
```

## Data Flow Diagrams

### 1. Presentation Creation Flow
```mermaid
sequenceDiagram
    participant Client
    participant Controller
    participant JsonParser
    participant PPTService
    participant FileSystem

    Client->>Controller: POST /create-from-json
    Controller->>JsonParser: Parse JSON content
    JsonParser->>JsonParser: Validate & transform
    JsonParser-->>Controller: PresentationContent
    Controller->>PPTService: CreatePresentationAsync()
    PPTService->>PPTService: Generate OpenXML
    PPTService->>FileSystem: Save .pptx file
    FileSystem-->>PPTService: File path
    PPTService-->>Controller: Success response
    Controller-->>Client: PresentationResponse
```

### 2. Image Upload Flow
```mermaid
sequenceDiagram
    participant Client
    participant Controller
    participant FileSystem
    participant ImageHelper

    Client->>Controller: POST /upload-image
    Controller->>Controller: Validate file type & size
    Controller->>FileSystem: Check if file exists
    alt File exists
        FileSystem-->>Controller: File info
        Controller-->>Client: Skip upload response
    else File doesn't exist
        Controller->>FileSystem: Save image file
        Controller->>ImageHelper: Extract dimensions
        ImageHelper-->>Controller: Image metadata
        Controller-->>Client: Upload success response
    end
```

### 3. System Startup Flow
```mermaid
graph LR
    START[Application Start] --> CONFIG[Load Configuration]
    CONFIG --> SERVICES[Register Services]
    SERVICES --> MIDDLEWARE[Configure Middleware]
    MIDDLEWARE --> ROUTES[Map Routes]
    ROUTES --> DIRS[Create Directories]
    DIRS --> SWAGGER[Configure Swagger]
    SWAGGER --> LISTEN[Start Listening]
    
    subgraph "Directory Setup"
        DIRS --> IMG_DIR[Images Directory]
        DIRS --> PPT_DIR[Presentations Directory]
    end
```

## Deployment Architecture Options

### Option 1: Single Server Deployment
```mermaid
graph TB
    subgraph "Single Server"
        subgraph "Application Server"
            API[ASP.NET Core API]
            FS[Local File System]
        end
        
        subgraph "Storage"
            IMG_STORE[Images Storage]
            PPT_STORE[Presentations Storage]
        end
    end
    
    CLIENT[Clients] --> API
    API --> FS
    FS --> IMG_STORE
    FS --> PPT_STORE
```

### Option 2: Cloud Deployment (Azure)
```mermaid
graph TB
    subgraph "Azure Cloud"
        subgraph "Compute"
            APP_SERVICE[Azure App Service]
        end
        
        subgraph "Storage"
            BLOB[Azure Blob Storage<br/>Images]
            FILES[Azure Files<br/>Presentations]
        end
        
        subgraph "Monitoring"
            INSIGHTS[Application Insights]
            LOGS[Azure Monitor]
        end
        
        subgraph "Security"
            KEYVAULT[Azure Key Vault]
            AAD[Azure AD]
        end
    end
    
    CLIENT[Clients] --> APP_SERVICE
    APP_SERVICE --> BLOB
    APP_SERVICE --> FILES
    APP_SERVICE --> INSIGHTS
    APP_SERVICE --> KEYVAULT
```

### Option 3: Containerized Deployment
```mermaid
graph TB
    subgraph "Container Orchestration"
        subgraph "Kubernetes/Docker"
            POD1[API Pod 1]
            POD2[API Pod 2]
            POD3[API Pod 3]
        end
        
        LB[Load Balancer]
        
        subgraph "Persistent Storage"
            PVC[Persistent Volume Claims]
            NFS[Network File System]
        end
    end
    
    CLIENT[Clients] --> LB
    LB --> POD1
    LB --> POD2
    LB --> POD3
    
    POD1 --> PVC
    POD2 --> PVC
    POD3 --> PVC
    PVC --> NFS
```

## Security Architecture

### Current Security Model
```mermaid
graph LR
    REQUEST[Client Request] --> CORS_CHECK[CORS Validation]
    CORS_CHECK --> FILE_VAL[File Validation]
    FILE_VAL --> SIZE_CHECK[Size Limits]
    SIZE_CHECK --> TYPE_CHECK[Type Validation]
    TYPE_CHECK --> PROCESS[Process Request]
    PROCESS --> SANITIZE[Sanitize Response]
    SANITIZE --> RESPONSE[Send Response]
```

### Future Security Enhancements
```mermaid
graph TB
    subgraph "Authentication Layer"
        JWT[JWT Tokens]
        API_KEY[API Keys]
        OAUTH[OAuth 2.0]
    end
    
    subgraph "Authorization Layer"
        RBAC[Role-Based Access]
        PERMISSIONS[Resource Permissions]
        RATE_LIMIT[Rate Limiting]
    end
    
    subgraph "Data Protection"
        ENCRYPT[Data Encryption]
        SANITIZE[Input Sanitization]
        AUDIT[Audit Logging]
    end
    
    CLIENT[Client] --> JWT
    JWT --> RBAC
    RBAC --> ENCRYPT
    ENCRYPT --> API_ENDPOINT[API Endpoints]
```

This comprehensive architecture documentation provides both high-level system design and detailed component relationships for your PowerPoint Generator web service.
