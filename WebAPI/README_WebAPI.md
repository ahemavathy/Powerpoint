# PowerPoint Generator Web API

A REST API service for generating PowerPoint presentations from JSON content using both programmatic generation and template-based approaches. This service can be called from any web application, mobile app, or system that supports HTTP requests.

## 🚀 Quick Start

### 1. Run the Web Service

```bash
# Navigate to the WebAPI directory
cd WebAPI

# Run the web API (development mode)
dotnet run --project PowerPointGenerator.WebAPI.csproj

# Or from the root directory
dotnet run --project WebAPI/PowerPointGenerator.WebAPI.csproj
```

The API will start at:
- **HTTP**: http://localhost:5000
- **HTTPS**: https://localhost:7000 (if configured)
- **Swagger UI**: http://localhost:5000 (in development mode)

### 2. Test the API

Once running, visit http://localhost:5000 to see the Swagger documentation and test the endpoints interactively.

## 📝 API Endpoints

### Create Presentation (Programmatic)
**POST** `/api/presentation/create-from-json`

Creates a PowerPoint presentation from JSON slide content using programmatic generation.

**Request Body:**
```json
{
  "jsonContent": "{\"slides\": [{\"title\": \"My Title\", \"description\": \"My description\", \"suggested_image\": \"image.png\"}]}",
  "presentationName": "MyPresentation",
  "presentationTitle": "My Presentation Title",
  "author": "John Doe"
}
```

### Create Presentation from Template (NEW!)
**POST** `/api/presentation/create-from-template`

Creates a PowerPoint presentation using an existing template with placeholder replacement. This provides more control over the exact layout, fonts, colors, and design.

**Request Body:**
```json
{
  "jsonContent": "{\"slides\": [{\"title\": \"My Title\", \"description\": \"My description\", \"suggested_image\": \"image.png\"}]}",
  "presentationName": "MyPresentation",
  "presentationTitle": "My Presentation Title", 
  "author": "John Doe",
  "templateName": "my_template.pptx"
}
```

**Template Features:**
- Uses PowerPoint templates from the `Templates/` folder
- Supports text placeholders: `{{TITLE}}`, `{{DESCRIPTION}}`, `{{SYNOPSIS}}`
- Replaces images automatically with your content images
- Maintains exact template formatting and design
- Removes excess slides if template has more slides than content
- Removes images from slides that don't have images in content
- Preserves aspect ratios of replacement images

**Available Templates:**
- Get list of templates: **GET** `/api/presentation/templates`
- Default template: `test_template.pptx` (used when no template specified)

**Response:**
```json
{
  "success": true,
  "fileName": "MyPresentation_abc123.pptx",
  "filePath": "/path/to/file.pptx",
  "presentationName": "MyPresentation",
  "createdAt": "2025-07-29T12:00:00Z",
  "fileSize": 45678,
  "slideCount": 3,
  "downloadUrl": "/api/presentation/download/MyPresentation_abc123.pptx"
}
```

### Image Upload
**POST** `/api/presentation/upload-image`

Uploads an image file for use in presentations.

**POST** `/api/presentation/upload-images`

Uploads multiple image files.

### File Management
- **GET** `/api/presentation/download/{fileName}` - Downloads a generated presentation file
- **GET** `/api/presentation/list` - Gets a list of all generated presentations
- **DELETE** `/api/presentation/delete/{fileName}` - Deletes a generated presentation file
- **GET** `/api/presentation/images` - Gets a list of uploaded images
- **GET** `/api/presentation/image/{fileName}` - Downloads/views an uploaded image
- **DELETE** `/api/presentation/image/{fileName}` - Deletes an uploaded image

### Health Check
**GET** `/api/presentation/health`

Checks if the API service is running.

## 🔧 Usage Examples

### JavaScript/TypeScript (Web Application)
```javascript
// Create presentation (programmatic)
const createPresentation = async () => {
  const jsonContent = {
    slides: [
      {
        title: "Welcome",
        description: "This is a test slide",
        suggested_image: "welcome.png"
      }
    ]
  };

  const response = await fetch('http://localhost:5000/api/presentation/create-from-json', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      jsonContent: JSON.stringify(jsonContent),
      presentationName: 'WebApp_Presentation',
      presentationTitle: 'My Web App Presentation',
      author: 'Web User'
    })
  });

  const result = await response.json();
  console.log('Presentation created:', result);

  // Download the file
  if (result.success) {
    window.open(`http://localhost:5000${result.downloadUrl}`);
  }
};

// Create presentation from template (NEW!)
const createFromTemplate = async () => {
  const jsonContent = {
    slides: [
      {
        title: "Product Launch",
        description: "Introducing our amazing new product",
        suggested_image: "product.png"
      },
      {
        title: "Key Features", 
        description: "Advanced functionality that sets us apart"
      }
    ]
  };

  const response = await fetch('http://localhost:5000/api/presentation/create-from-template', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      jsonContent: JSON.stringify(jsonContent),
      presentationName: 'Template_Presentation',
      presentationTitle: 'Professional Template Presentation',
      author: 'Marketing Team',
      templateName: 'corporate_template.pptx' // optional, uses test_template.pptx if not specified
    })
  });

  const result = await response.json();
  console.log('Template presentation created:', result);
};

// Upload image
const uploadImage = async (file) => {
  const formData = new FormData();
  formData.append('file', file);

  const response = await fetch('http://localhost:5000/api/presentation/upload-image', {
    method: 'POST',
    body: formData
  });

  const result = await response.json();
  console.log('Image uploaded:', result);
};
```

### Python
```python
import requests
import json

def create_presentation():
    json_content = {
        "slides": [
            {
                "title": "Python Generated Slide",
                "description": "This slide was created from Python",
                "suggested_image": "python_logo.png"
            }
        ]
    }
    
    payload = {
        "jsonContent": json.dumps(json_content),
        "presentationName": "Python_Presentation",
        "presentationTitle": "Generated from Python",
        "author": "Python Script"
    }
    
    response = requests.post(
        'http://localhost:5000/api/presentation/create-from-json',
        json=payload
    )
    
    if response.ok:
        result = response.json()
        print(f"Presentation created: {result['fileName']}")
        
        # Download the file
        download_url = f"http://localhost:5000{result['downloadUrl']}"
        file_response = requests.get(download_url)
        
        with open(result['fileName'], 'wb') as f:
            f.write(file_response.content)
        print(f"Downloaded: {result['fileName']}")
    else:
        print(f"Error: {response.text}")

def upload_image(file_path):
    with open(file_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(
            'http://localhost:5000/api/presentation/upload-image',
            files=files
        )
    
    if response.ok:
        result = response.json()
        print(f"Image uploaded: {result['fileName']}")
    else:
        print(f"Upload error: {response.text}")

# Usage
create_presentation()
upload_image('path/to/your/image.png')
```

### C# (.NET)
```csharp
using System.Text;
using System.Text.Json;

public class PowerPointApiClient
{
    private readonly HttpClient _httpClient;
    private readonly string _baseUrl;

    public PowerPointApiClient(string baseUrl = "http://localhost:5000")
    {
        _baseUrl = baseUrl;
        _httpClient = new HttpClient();
    }

    public async Task<string> CreatePresentationAsync(string jsonContent, string presentationName)
    {
        var payload = new
        {
            jsonContent = jsonContent,
            presentationName = presentationName,
            presentationTitle = presentationName,
            author = "API Client"
        };

        var json = JsonSerializer.Serialize(payload);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        var response = await _httpClient.PostAsync($"{_baseUrl}/api/presentation/create-from-json", content);
        return await response.Content.ReadAsStringAsync();
    }

    public async Task<byte[]> DownloadPresentationAsync(string fileName)
    {
        var response = await _httpClient.GetAsync($"{_baseUrl}/api/presentation/download/{fileName}");
        return await response.Content.ReadAsByteArrayAsync();
    }
}

// Usage
var client = new PowerPointApiClient();
var jsonContent = @"{""slides"": [{""title"": ""C# Generated Slide"", ""description"": ""Created from C#"", ""suggested_image"": ""logo.png""}]}";
var result = await client.CreatePresentationAsync(jsonContent, "CSharp_Presentation");
```

### cURL (Command Line)
```bash
# Create presentation (programmatic)
curl -X POST "http://localhost:5000/api/presentation/create-from-json" \
  -H "Content-Type: application/json" \
  -d '{
    "jsonContent": "{\"slides\": [{\"title\": \"cURL Test\", \"description\": \"Created via cURL\", \"suggested_image\": \"test.png\"}]}",
    "presentationName": "cURL_Test",
    "presentationTitle": "cURL Generated Presentation",
    "author": "Command Line User"
  }'

# Create presentation from template (NEW!)
curl -X POST "http://localhost:5000/api/presentation/create-from-template" \
  -H "Content-Type: application/json" \
  -d '{
    "jsonContent": "{\"slides\": [{\"title\": \"Template Test\", \"description\": \"Created via template\", \"suggested_image\": \"logo.png\"}]}",
    "presentationName": "Template_Test",
    "presentationTitle": "Template Generated Presentation",
    "author": "Template User",
    "templateName": "test_template.pptx"
  }'

# List available templates
curl "http://localhost:5000/api/presentation/templates"

# Upload image
curl -X POST "http://localhost:5000/api/presentation/upload-image" \
  -F "file=@path/to/your/image.png"

# Download presentation (replace FILENAME with actual filename from response)
curl -O "http://localhost:5000/api/presentation/download/FILENAME.pptx"

# List presentations
curl "http://localhost:5000/api/presentation/list"

# Health check
curl "http://localhost:5000/api/presentation/health"
```

## 🔒 Configuration

### CORS
The API is configured to accept requests from any origin in development mode. For production, configure CORS appropriately in the startup configuration.

### File Storage
- Generated presentations are stored in `WebAPI/GeneratedPresentations/` directory
- Uploaded images are stored in `WebAPI/Images/` directory
- PowerPoint templates are stored in `WebAPI/Templates/` directory
- All directories are created automatically when the service starts

### Logging
The API includes comprehensive logging for debugging and monitoring.

## 🏗️ Project Structure
```
PowerPointGenerator/
├── WebAPI/
│   ├── Controllers/
│   │   └── PresentationController.cs    # Main API controller
│   ├── Models/
│   │   └── WebApiModels.cs              # API request/response models
│   ├── Properties/
│   │   └── launchSettings.json          # Launch configuration
│   ├── Program.cs                       # Web API startup
│   ├── PowerPointGenerator.WebAPI.csproj
│   ├── GeneratedPresentations/          # Output directory
│   ├── Images/                          # Uploaded images directory
│   └── Templates/                       # PowerPoint template files (NEW!)
│       └── test_template.pptx          # Default template
├── Controllers/                         # Shared controllers
├── Models/                             # Core domain models
├── Services/                           # Core presentation generation logic
├── Utilities/                          # Helper classes
└── Client/                             # C# client library
```

## 🚀 Deployment

### Development
```bash
# From WebAPI directory
cd WebAPI
dotnet run

# Or from root directory
dotnet run --project WebAPI/PowerPointGenerator.WebAPI.csproj
```

### Production
```bash
# Build for production
dotnet publish WebAPI/PowerPointGenerator.WebAPI.csproj -c Release -o ./publish

# Run published version
cd publish
dotnet PowerPointGenerator.WebAPI.dll
```

### Docker
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY ./publish .
EXPOSE 80
ENTRYPOINT ["dotnet", "PowerPointGenerator.WebAPI.dll"]
```

## 🧪 Testing

### Interactive Testing
- **Swagger UI**: http://localhost:5000 (interactive API documentation)
- **Health Check**: http://localhost:5000/api/presentation/health

### Automated Testing
- Unit tests can be added in a separate test project
- Integration tests can use the provided client examples

## 📋 Requirements

- .NET 8.0 SDK
- Windows (for System.Drawing.Common image processing)
- Sufficient disk space for generated presentations and uploaded images

## 🔍 Troubleshooting

**Common Issues:**
1. **Port already in use**: Change the port in `WebAPI/Properties/launchSettings.json`
2. **Permission errors**: Ensure the service has write permissions to the output directories
3. **Image files not found**: Upload images via the API or place them in the `WebAPI/Images/` directory
4. **CORS errors**: Configure CORS settings for your specific domain in production
5. **Build errors**: Ensure .NET 8.0 SDK is installed and all NuGet packages are restored

**Debug Steps:**
1. Check console output for detailed error messages
2. Verify the health endpoint is responding
3. Check file permissions on output directories
4. Validate JSON format in requests

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is available under the MIT License. See the LICENSE file for more details.

---

**Ready to generate PowerPoint presentations from any application! 🎉**
