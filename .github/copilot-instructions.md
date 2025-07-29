# Copilot Instructions for PowerPoint Generator

<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

This is a C# .NET project that uses the Open-XML-SDK to create PowerPoint presentations with image-heavy slides.

## Project Context
- **Language**: C#
- **Framework**: .NET 8.0
- **Main Library**: DocumentFormat.OpenXml (Open-XML-SDK)
- **Purpose**: Generate PowerPoint presentations from AI-generated content including images and synopsis

## Key Guidelines
1. Use DocumentFormat.OpenXml namespace for PowerPoint manipulation
2. Focus on image-heavy slide creation capabilities
3. Handle AI-generated input (images + synopsis text)
4. Follow modern C# patterns and best practices
5. Implement robust error handling for file operations
6. Use async/await patterns for I/O operations where appropriate

## Project Structure
- `Models/`: Data models for presentation content
- `Services/`: Core business logic for PowerPoint generation
- `Utilities/`: Helper classes and extensions
- `Examples/`: Sample usage and test data

## Code Style
- Use PascalCase for public members
- Use camelCase for private fields and local variables
- Implement IDisposable pattern for OpenXml document handling
- Add comprehensive XML documentation for public APIs
