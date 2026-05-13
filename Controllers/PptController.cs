using Microsoft.AspNetCore.Mvc;
using ppt_mapper.Models;
using ppt_mapper.Services;

[ApiController]
[Route("api/ppt")]
public class PptController : ControllerBase
{
    private readonly PptService _pptService;

    public PptController(PptService pptService)
    {
        _pptService = pptService;
    }

    [HttpPost("generate")]
    [Consumes("multipart/form-data")]
    public IActionResult Generate([FromForm] PptRequest request)
    {
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "template.pptx");
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), $"output_{Guid.NewGuid()}.pptx");

        System.IO.File.Copy(templatePath, outputPath, true);

        var data = new Dictionary<string, string>
    {
        { "{{TITLE}}", request.Title },
        { "{{DATE}}", request.Date },
        { "{{N}}", request.SprintNumber },
        { "{{POINTS}}", request.Points != null && request.Points.Any()
            ? "• " + string.Join("\n• ", request.Points)
            : "" }
    };

        _pptService.ReplacePlaceholders(outputPath, data);

        // 🔥 Save uploaded images temporarily
        var imagePaths = new List<string>();

        if (request.Images != null && request.Images.Any())
        {
            foreach (var file in request.Images)
            {
                var tempPath = Path.GetTempFileName();

                using (var stream = new FileStream(tempPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                imagePaths.Add(tempPath);
            }

            _pptService.ReplaceImages(outputPath, imagePaths);
        }

        var bytes = System.IO.File.ReadAllBytes(outputPath);

        return File(bytes,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "Generated.pptx");
    }
}