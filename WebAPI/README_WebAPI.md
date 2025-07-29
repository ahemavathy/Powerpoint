# PowerPoint Generator Web API

A REST API service for generating PowerPoint presentations from JSON content. This service can be called from any web application, mobile app, or system that supports HTTP requests.

## 🚀 Quick Start

### 1. Run the Web Service

```bash
# Navigate to the project directory
cd c:\Users\hema\Projects\Powerpoint

# Run the web API (development mode)
dotnet run --project PowerPointGenerator.WebAPI.csproj --launch-profile http

# Or run directly with the web program
dotnet run WebProgram.cs
```

The API will start at:
- **HTTP**: http://localhost:5000
- **HTTPS**: https://localhost:7000 (if configured)
- **Swagger UI**: http://localhost:5000 (in development mode)

### 2. Test the API

Once running, visit http://localhost:5000 to see the Swagger documentation and test the endpoints interactively.

## 📝 API Endpoints

### Create Presentation
**POST** `/api/presentation/create-from-json`

Creates a PowerPoint presentation from JSON slide content.

**Request Body:**
```json
{
  "jsonContent": "{\"slides\": [{\"title\": \"My Title\", \"description\": \"My description\", \"suggested_image\": \"image.png\"}]}",
  "presentationName": "MyPresentation",
  "presentationTitle": "My Presentation Title",
  "author": "John Doe"
}
```

**Response:**
```json
{
  "success": true,
  "fileName": "MyPresentation_abc123.pptx",
  "filePath": "/path/to/file.pptx",
  "presentationName": "MyPresentation",
  "createdAt": "2025-07-22T12:00:00Z",
  "fileSize": 45678,
  "slideCount": 3,
  "downloadUrl": "/api/presentation/download/MyPresentation_abc123.pptx"
}
```

### Download Presentation
**GET** `/api/presentation/download/{fileName}`

Downloads a generated presentation file.

### List Presentations
**GET** `/api/presentation/list`

Gets a list of all generated presentations.

### Health Check
**GET** `/api/presentation/health`

Checks if the API service is running.

### Delete Presentation
**DELETE** `/api/presentation/delete/{fileName}`

Deletes a generated presentation file.

## 🔧 Usage Examples

### JavaScript/TypeScript (Web Application)
```javascript
// Create presentation
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

create_presentation()
```

### C# (.NET)
```csharp
using PowerPointGenerator.Client;

// Using the provided client
var client = new PowerPointApiClient("http://localhost:5000");

var jsonContent = @"{
  ""slides"": [
    {
      ""title"": ""C# Generated Slide"",
      ""description"": ""This was created from a C# application"",
      ""suggested_image"": ""csharp_logo.png""
    }
  ]
}";

var response = await client.CreatePresentationAsync(
    jsonContent,
    "CSharp_Presentation",
    "Generated from C#",
    "C# Developer"
);

if (response.Success)
{
    var fileData = await client.DownloadPresentationAsync(response.FileName);
    await File.WriteAllBytesAsync($"Downloaded_{response.FileName}", fileData);
}
```

### cURL (Command Line)
```bash
# Create presentation
curl -X POST "http://localhost:5000/api/presentation/create-from-json" \
  -H "Content-Type: application/json" \
  -d '{
    "jsonContent": "{\"slides\": [{\"title\": \"cURL Test\", \"description\": \"Created via cURL\", \"suggested_image\": \"test.png\"}]}",
    "presentationName": "cURL_Test",
    "presentationTitle": "cURL Generated Presentation",
    "author": "Command Line User"
  }'

# Download presentation (replace FILENAME with actual filename from response)
curl -O "http://localhost:5000/api/presentation/download/FILENAME.pptx"

# List presentations
curl "http://localhost:5000/api/presentation/list"

# Health check
curl "http://localhost:5000/api/presentation/health"
```

## 🔒 Configuration

### CORS
The API is configured to accept requests from any origin in development mode. For production, configure CORS appropriately in `WebProgram.cs`.

### File Storage
- Generated presentations are stored in `GeneratedPresentations/` directory
- Image files should be placed in `Images/` directory
- Both directories are created automatically when the service starts

### Logging
The API includes comprehensive logging for debugging and monitoring.

## 🏗️ Project Structure
```
PowerPointGenerator/
├── Controllers/
│   └── PresentationController.cs    # Main API controller
├── Client/
│   └── PowerPointApiClient.cs       # C# client library
├── Examples/
│   └── ApiClientExample.cs         # Usage examples
├── Models/
│   └── WebApiModels.cs              # API request/response models
├── Services/                        # Core presentation generation logic
├── WebProgram.cs                    # Web API startup
└── PowerPointGenerator.WebAPI.csproj
```

## 🚀 Deployment

### Development
```bash
dotnet run --project PowerPointGenerator.WebAPI.csproj
```

### Production
```bash
# Build for production
dotnet publish -c Release -o ./publish

# Run published version
cd publish
dotnet PowerPointGenerator.WebAPI.dll
```

### Docker (Optional)
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY ./publish .
EXPOSE 80
ENTRYPOINT ["dotnet", "PowerPointGenerator.WebAPI.dll"]
```

## 🧪 Testing
- **Swagger UI**: http://localhost:5000 (interactive API documentation)
- **Health Check**: http://localhost:5000/api/presentation/health
- **Unit Tests**: Can be added in a separate test project

## 🔍 Troubleshooting

**Common Issues:**
1. **Port already in use**: Change the port in `launchSettings.json`
2. **Permission errors**: Ensure the service has write permissions to the output directories
3. **Image files not found**: Place images in the `Images/` directory or use absolute paths
4. **CORS errors**: Configure CORS settings for your specific domain in production

**Logs**: Check console output for detailed error messages and request logging.

Now you can call the PowerPoint Generator from any web application or system that supports HTTP requests! 🎉
