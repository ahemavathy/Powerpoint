# PowerPoint Generator - System Architecture Diagrams

This document provides visual representations of the PowerPoint Generator system architecture using Mermaid diagrams.

## System Overview Diagram

```mermaid
graph TB
    subgraph "Client Applications"
        CLI[Console Application]
        WEB[Web Applications]
        MOB[Mobile Apps]
        API[API Clients]
    end

    subgraph "PowerPoint Generator System"
        subgraph "Console App"
            PROG[Program.cs]
            PPAPI[PowerPointAPI.cs]
        end
        
        subgraph "Web API Layer"
            WEBAPI[ASP.NET Core 8.0]
            CTRL[PresentationController]
            SWAGGER[Swagger UI]
            CORS[CORS Middleware]
        end
        
        subgraph "Business Logic Layer"
            subgraph "Services"
                PPG[PowerPointGeneratorService]
                JSON[JsonSlideParser]
                SCP[SlideContentParser]
            end
            
            subgraph "Utilities"
                SH[SlideHelper]
                IH[ImageHelper]
                TH[ThemeHelper]
            end
        end
        
        subgraph "Models Layer"
            PM[PresentationModels]
            JSM[JsonSlideModels]
            WAM[WebApiModels]
        end
        
        subgraph "Data Access Layer"
            FS[File System]
            OXSDK[OpenXML SDK]
        end
        
        subgraph "Storage"
            IMG[Images/]
            PPTX[Generated Presentations/]
            JSON_FILES[JSON Input Files]
        end
    end

    CLI --> PROG
    CLI --> PPAPI
    WEB --> WEBAPI
    MOB --> WEBAPI
    API --> WEBAPI
    
    PROG --> JSON
    PPAPI --> JSON
    CTRL --> JSON
    
    JSON --> PPG
    SCP --> PPG
    PPG --> SH
    PPG --> IH
    PPG --> TH
    
    PPG --> OXSDK
    PPG --> FS
    
    FS --> IMG
    FS --> PPTX
    FS --> JSON_FILES
    
    PM -.-> PPG
    JSM -.-> JSON
    WAM -.-> CTRL
```

## Detailed API Data Flow

```mermaid
sequenceDiagram
    participant Client
    participant Controller
    participant JsonParser
    participant Generator
    participant FileSystem
    participant OpenXML

    Note over Client,OpenXML: Presentation Creation Flow
    
    Client->>Controller: POST /api/presentation/create-from-json
    Controller->>Controller: Validate Request
    Controller->>JsonParser: ParseFromString(jsonContent)
    JsonParser->>JsonParser: Parse JSON Structure
    JsonParser->>FileSystem: Check Image Files
    JsonParser-->>Controller: PresentationContent Model
    
    Controller->>Generator: CreatePresentationAsync()
    Generator->>Generator: Process Slides
    Generator->>OpenXML: Create PowerPoint Document
    Generator->>FileSystem: Read Image Files
    Generator->>OpenXML: Embed Images with Aspect Ratio
    Generator->>OpenXML: Apply Formatting & Themes
    Generator->>FileSystem: Save .pptx File
    Generator-->>Controller: File Path
    
    Controller->>Controller: Build Response
    Controller-->>Client: PresentationResponse (with download URL)

    Note over Client,OpenXML: Image Upload Flow
    
    Client->>Controller: POST /api/presentation/upload-image
    Controller->>Controller: Validate File Type & Size
    Controller->>FileSystem: Check if File Exists
    alt File Exists
        Controller-->>Client: Skip Upload (File Exists)
    else New File
        Controller->>FileSystem: Save Image File
        Controller-->>Client: Upload Success Response
    end

    Note over Client,OpenXML: File Download Flow
    
    Client->>Controller: GET /api/presentation/download/{fileName}
    Controller->>FileSystem: Read File
    Controller-->>Client: Binary File Stream
```

## Console Application Flow

```mermaid
flowchart TD
    START([Start Console App])
    ARGS{Parse Command Line Args}
    
    ARGS -->|No Args| DEFAULT[Use slides_content.json]
    ARGS -->|JSON File Only| FILENAME[Use JSON filename for output]
    ARGS -->|JSON + Name| CUSTOM[Use custom presentation name]
    
    DEFAULT --> PARSE[JsonSlideParser.ParseFromFile]
    FILENAME --> PARSE
    CUSTOM --> PARSE
    
    PARSE --> VALIDATE{Validate JSON}
    VALIDATE -->|Invalid| ERROR[Show Error & Exit]
    VALIDATE -->|Valid| GENERATE[PowerPointGeneratorService.CreatePresentationAsync]
    
    GENERATE --> IMAGES{Check Images}
    IMAGES -->|Missing| PLACEHOLDER[Create Placeholder Images]
    IMAGES -->|Found| EMBED[Embed Actual Images]
    
    PLACEHOLDER --> BUILD[Build Presentation]
    EMBED --> BUILD
    
    BUILD --> SAVE[Save .pptx File]
    SAVE --> SUCCESS[Show Success Message]
    SUCCESS --> END([End])
    ERROR --> END
```

## Current File Structure

```mermaid
graph LR
    subgraph "Project Root"
        subgraph "Console App Files"
            PROG_CS[Program.cs]
            PPAPI_CS[PowerPointAPI.cs]
            PROJ[PowerPointGenerator.csproj]
            JSON_SAMPLE[slides_content.json]
        end
        
        subgraph "Shared Components"
            MODELS[Models/]
            SERVICES[Services/]
            UTILITIES[Utilities/]
            CONTROLLERS[Controllers/]
            IMAGES[Images/]
        end
        
        subgraph "WebAPI Project"
            WEBAPI_DIR[WebAPI/]
            WEBAPI_PROG[WebAPI/Program.cs]
            WEBAPI_PROJ[WebAPI/PowerPointGenerator.WebAPI.csproj]
            WEBAPI_IMGS[WebAPI/Images/]
            WEBAPI_GEN[WebAPI/GeneratedPresentations/]
        end
    end
```

## Deployment Architecture Options

### Option 1: Single Server Deployment

```mermaid
graph TB
    subgraph "Production Server"
        subgraph "Application Runtime"
            NET8[.NET 8.0 Runtime]
            WEBAPP[PowerPoint Generator Web API]
            CONSOLE[Console Application]
        end
        
        subgraph "File Storage"
            UPLOAD_IMGS[Uploaded Images]
            GEN_PPTS[Generated Presentations]
            LOGS[Application Logs]
        end
        
        subgraph "Network"
            HTTP[HTTP :5000]
            HTTPS[HTTPS :7000]
        end
    end
    
    subgraph "External Clients"
        BROWSER[Web Browser]
        MOBILE[Mobile App]
        API_CLIENT[API Client]
        CLI_USER[Console User]
    end
    
    BROWSER --> HTTP
    MOBILE --> HTTPS
    API_CLIENT --> HTTP
    CLI_USER --> CONSOLE
    
    HTTP --> WEBAPP
    HTTPS --> WEBAPP
    WEBAPP --> UPLOAD_IMGS
    WEBAPP --> GEN_PPTS
    CONSOLE --> GEN_PPTS
```

### Option 2: Containerized Deployment

```mermaid
graph TB
    subgraph "Container Platform (Docker/Kubernetes)"
        subgraph "Web API Container"
            API_CONTAINER[PowerPoint Generator API]
            API_PORT[":80"]
        end
        
        subgraph "Volumes"
            VOLUME_IMGS[Images Volume]
            VOLUME_PPTS[Presentations Volume]
        end
        
        subgraph "Console Container"
            CONSOLE_CONTAINER[Console App Container]
        end
    end
    
    subgraph "External Storage"
        BLOB[Azure Blob Storage / AWS S3]
        NFS[Network File System]
    end
    
    subgraph "Load Balancer"
        LB[Load Balancer]
    end
    
    LB --> API_CONTAINER
    API_CONTAINER --> VOLUME_IMGS
    API_CONTAINER --> VOLUME_PPTS
    CONSOLE_CONTAINER --> VOLUME_PPTS
    
    VOLUME_IMGS -.-> BLOB
    VOLUME_PPTS -.-> NFS
```

### Option 3: Cloud Native Architecture

```mermaid
graph TB
    subgraph "Azure/AWS Cloud"
        subgraph "Compute"
            APP_SERVICE[App Service / ECS]
            FUNCTION[Azure Functions / Lambda]
        end
        
        subgraph "Storage"
            BLOB_STORAGE[Blob Storage / S3]
            FILE_SHARE[File Share / EFS]
        end
        
        subgraph "Networking"
            CDN[CDN]
            API_GATEWAY[API Gateway]
        end
        
        subgraph "Monitoring"
            APP_INSIGHTS[Application Insights]
            CLOUDWATCH[CloudWatch]
        end
    end
    
    subgraph "Clients"
        WEB_CLIENT[Web Clients]
        MOBILE_CLIENT[Mobile Clients]
        API_CLIENTS[API Clients]
    end
    
    WEB_CLIENT --> CDN
    MOBILE_CLIENT --> API_GATEWAY
    API_CLIENTS --> API_GATEWAY
    
    CDN --> APP_SERVICE
    API_GATEWAY --> APP_SERVICE
    
    APP_SERVICE --> BLOB_STORAGE
    APP_SERVICE --> FILE_SHARE
    FUNCTION --> BLOB_STORAGE
    
    APP_SERVICE --> APP_INSIGHTS
    APP_SERVICE --> CLOUDWATCH
```

## Security Architecture

### Current Security Model

```mermaid
graph TB
    subgraph "Client Layer"
        CLIENT[Client Applications]
    end
    
    subgraph "API Gateway Layer"
        CORS[CORS Policy]
        RATE_LIMIT[Rate Limiting]
        INPUT_VAL[Input Validation]
    end
    
    subgraph "Application Layer"
        AUTH[Authentication*]
        AUTHZ[Authorization*]
        FILE_VAL[File Validation]
        ERROR_HAND[Error Handling]
    end
    
    subgraph "Data Layer"
        FILE_SYS[File System Access]
        PATH_VAL[Path Validation]
        SIZE_LIMIT[Size Limits]
    end
    
    CLIENT --> CORS
    CORS --> RATE_LIMIT
    RATE_LIMIT --> INPUT_VAL
    INPUT_VAL --> AUTH
    AUTH --> AUTHZ
    AUTHZ --> FILE_VAL
    FILE_VAL --> ERROR_HAND
    ERROR_HAND --> FILE_SYS
    FILE_SYS --> PATH_VAL
    PATH_VAL --> SIZE_LIMIT
    
    note1[*Future Enhancement]
    AUTH -.-> note1
    AUTHZ -.-> note1
```

### Future Security Architecture

```mermaid
graph TB
    subgraph "Identity Provider"
        IDP[Azure AD / Auth0]
        JWT[JWT Token Service]
    end
    
    subgraph "API Gateway"
        TOKEN_VAL[Token Validation]
        RBAC[Role-Based Access Control]
        API_KEY[API Key Management]
    end
    
    subgraph "Application Security"
        ENCRYPT[TLS 1.2+ Encryption]
        AUDIT[Audit Logging]
        SANITIZE[Input Sanitization]
    end
    
    subgraph "Data Security"
        ENCRYPT_REST[Encryption at Rest]
        ACCESS_CTRL[File Access Control]
        BACKUP[Secure Backup]
    end
    
    IDP --> JWT
    JWT --> TOKEN_VAL
    TOKEN_VAL --> RBAC
    RBAC --> API_KEY
    API_KEY --> ENCRYPT
    ENCRYPT --> AUDIT
    AUDIT --> SANITIZE
    SANITIZE --> ENCRYPT_REST
    ENCRYPT_REST --> ACCESS_CTRL
    ACCESS_CTRL --> BACKUP
```

## Technology Stack Diagram

```mermaid
graph TB
    subgraph "Runtime Environment"
        NET8[.NET 8.0]
        ASPNET[ASP.NET Core 8.0]
    end
    
    subgraph "Core Libraries"
        OPENXML[DocumentFormat.OpenXml 3.3.0/3.0.1]
        DRAWING[System.Drawing.Common 9.0.7/8.0.0]
        JSON_LIB[System.Text.Json 8.0.0]
    end
    
    subgraph "Web API Libraries"
        SWAGGER[Swashbuckle.AspNetCore 6.4.0]
        OPENAPI[Microsoft.AspNetCore.OpenApi 8.0.0]
    end
    
    subgraph "Development Tools"
        VSCODE[VS Code]
        DOTNET_CLI[.NET CLI]
        GIT[Git]
    end
    
    subgraph "File Formats"
        PPTX[PowerPoint .pptx]
        JSON_FORMAT[JSON Input]
        IMAGE_FORMATS[JPG/PNG/GIF/BMP/WEBP]
    end
    
    NET8 --> ASPNET
    ASPNET --> OPENXML
    ASPNET --> DRAWING
    ASPNET --> JSON_LIB
    ASPNET --> SWAGGER
    ASPNET --> OPENAPI
    
    OPENXML --> PPTX
    JSON_LIB --> JSON_FORMAT
    DRAWING --> IMAGE_FORMATS
```

## System Startup Flow

```mermaid
sequenceDiagram
    participant User
    participant Runtime
    participant WebAPI
    participant FileSystem
    participant Services

    Note over User,Services: Web API Startup
    
    User->>Runtime: dotnet run WebAPI/PowerPointGenerator.WebAPI.csproj
    Runtime->>WebAPI: Initialize Application
    WebAPI->>WebAPI: Configure Services
    WebAPI->>WebAPI: Configure Middleware Pipeline
    WebAPI->>FileSystem: Create Required Directories
    FileSystem-->>WebAPI: Directories Ready
    WebAPI->>Services: Register Dependencies
    Services-->>WebAPI: Services Registered
    WebAPI->>WebAPI: Configure Swagger UI
    WebAPI->>Runtime: Start HTTP Server
    Runtime-->>User: API Ready (http://localhost:5000)

    Note over User,Services: Console App Startup
    
    User->>Runtime: dotnet run --project PowerPointGenerator.csproj
    Runtime->>Services: Initialize JsonSlideParser
    Runtime->>Services: Initialize PowerPointGeneratorService
    Services->>FileSystem: Check Image Directory
    FileSystem-->>Services: Directory Status
    Services->>Services: Process Input JSON
    Services->>Services: Generate Presentation
    Services->>FileSystem: Save Output File
    FileSystem-->>Services: File Saved
    Services-->>User: Success Message + File Location
```

This comprehensive set of diagrams provides visual documentation for all aspects of the PowerPoint Generator system architecture, from high-level system overview to detailed deployment scenarios and security models.
