using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Reflection;

namespace PowerPointGenerator.WebAPI.Filters
{
    /// <summary>
    /// Swagger operation filter to properly handle file upload parameters
    /// </summary>
    public class FileUploadOperationFilter : IOperationFilter
    {
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            var fileParams = context.MethodInfo.GetParameters()
                .Where(p => p.ParameterType == typeof(IFormFile) || 
                           p.ParameterType == typeof(List<IFormFile>) ||
                           p.ParameterType == typeof(IEnumerable<IFormFile>) ||
                           p.ParameterType == typeof(IFormFileCollection))
                .ToList();

            if (!fileParams.Any())
                return;

            operation.RequestBody = new OpenApiRequestBody
            {
                Content = new Dictionary<string, OpenApiMediaType>
                {
                    ["multipart/form-data"] = new OpenApiMediaType
                    {
                        Schema = new OpenApiSchema
                        {
                            Type = "object",
                            Properties = new Dictionary<string, OpenApiSchema>()
                        }
                    }
                }
            };

            foreach (var param in fileParams)
            {
                var schema = param.ParameterType == typeof(IFormFile)
                    ? new OpenApiSchema
                    {
                        Type = "string",
                        Format = "binary"
                    }
                    : new OpenApiSchema
                    {
                        Type = "array",
                        Items = new OpenApiSchema
                        {
                            Type = "string",
                            Format = "binary"
                        }
                    };

                operation.RequestBody.Content["multipart/form-data"].Schema.Properties.Add(
                    param.Name ?? "file", schema);
            }

            // Remove the file parameter from the parameters list since it's now in the request body
            var parametersToRemove = operation.Parameters
                .Where(p => fileParams.Any(fp => fp.Name == p.Name))
                .ToList();

            foreach (var param in parametersToRemove)
            {
                operation.Parameters.Remove(param);
            }
        }
    }
}
