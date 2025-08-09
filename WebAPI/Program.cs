using PowerPointGenerator.Services;
using PowerPointGenerator.WebAPI.Filters;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddScoped<PowerPointGeneratorService>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new() 
    { 
        Title = "PowerPoint Generator API", 
        Version = "v1",
        Description = "A Web API for generating PowerPoint presentations from JSON content",
        Contact = new()
        {
            Name = "PowerPoint Generator",
            Email = "support@powerpointgenerator.com"
        }
    });
    
    // Include XML comments for better documentation
    var xmlFile = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
    {
        c.IncludeXmlComments(xmlPath);
    }
    
    // Configure for file upload support
    c.OperationFilter<FileUploadOperationFilter>();
});

// Configure CORS for cross-origin requests
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// Add logging
builder.Services.AddLogging(logging =>
{
    logging.AddConsole();
    logging.AddDebug();
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "PowerPoint Generator API V1");
        c.RoutePrefix = string.Empty; // Make Swagger UI the root page
    });
}

app.UseHttpsRedirection();
app.UseCors();
app.UseAuthorization();
app.MapControllers();

// Ensure required directories exist
var requiredDirectories = new[]
{
    "GeneratedPresentations",
    "Images"
};

foreach (var dir in requiredDirectories)
{
    var path = Path.Combine(Environment.CurrentDirectory, dir);
    if (!Directory.Exists(path))
    {
        Directory.CreateDirectory(path);
        Console.WriteLine($"Created directory: {path}");
    }
}

Console.WriteLine("PowerPoint Generator Web API is starting...");
Console.WriteLine($"Environment: {app.Environment.EnvironmentName}");
Console.WriteLine("Available endpoints:");
Console.WriteLine("- Swagger UI: http://localhost:5000 (development)");
Console.WriteLine("- Health Check: GET /api/presentation/health");
Console.WriteLine("- Create Presentation: POST /api/presentation/create-from-json");
Console.WriteLine("- Download File: GET /api/presentation/download/{fileName}");
Console.WriteLine("- List Presentations: GET /api/presentation/list");

app.Run();
