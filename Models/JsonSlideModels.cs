using System.Text.Json.Serialization;

namespace PowerPointGenerator.Models
{
    /// <summary>
    /// Root JSON model for slide content
    /// </summary>
    public class JsonSlideContent
    {
        [JsonPropertyName("slides")]
        public List<JsonSlide> Slides { get; set; } = new List<JsonSlide>();
    }

    /// <summary>
    /// Individual slide in JSON format
    /// </summary>
    public class JsonSlide
    {
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        [JsonPropertyName("suggested_image")]
        public string SuggestedImage { get; set; } = string.Empty;

        [JsonPropertyName("layout")]
        public string Layout { get; set; } = string.Empty;
    }
}
