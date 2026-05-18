using Microsoft.AspNetCore.Mvc;
using ppt_mapper.Models;
using ppt_mapper.Services;

[ApiController]
[Route("api/ppt")]
public class PptController : ControllerBase
{
    private static readonly int[] MatrixNumericTokenOrder =
    [
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21
    ];

    private readonly PptService _pptService;

    public PptController(PptService pptService)
    {
        _pptService = pptService;
    }

    [HttpPost("generate")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> Generate([FromForm] PptRequest request)
    {
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "template.pptx");
        if (!System.IO.File.Exists(templatePath))
        {
            return NotFound("PPT template file was not found.");
        }

        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), $"output_{Guid.NewGuid():N}.pptx");
        var tempImagePaths = new List<string>();

        try
        {
            System.IO.File.Copy(templatePath, outputPath, true);

            var placeholderData = BuildPlaceholderData(request);
            _pptService.ReplacePlaceholders(outputPath, placeholderData);

            var imageMappings = await SaveImagesAsync(request.Images, tempImagePaths);
            _pptService.ReplaceImages(outputPath, imageMappings);

            var bytes = await System.IO.File.ReadAllBytesAsync(outputPath);
            var downloadFileName = BuildDownloadFileName(request.Date);

            return File(
                bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                downloadFileName);
        }
        finally
        {
            foreach (var imagePath in tempImagePaths)
            {
                TryDeleteFile(imagePath);
            }

            TryDeleteFile(outputPath);
        }
    }

    private static Dictionary<string, PptTextValue> BuildPlaceholderData(PptRequest request)
    {
        var data = new Dictionary<string, PptTextValue>(StringComparer.Ordinal);

        AddReplacement(data, "{{TITLE}}", request.Title, request.TitleIsBold);
        AddReplacement(data, "{{DATE}}", request.Date, request.DateIsBold);
        AddReplacement(data, "{{N}}", request.SprintNumber, request.SprintNumberIsBold);

        var legacyPointQueue = new Queue<string>(
            (request.Points ?? [])
                .Select(Normalize)
                .Where(static value => !string.IsNullOrWhiteSpace(value)));

        AddPointPage(data, 1, 11, request.Page1Points, request.Page1PointIsBold, request.Page1PointIsBulletPoint, legacyPointQueue);
        AddPointPage(data, 2, 11, request.Page2Points, request.Page2PointIsBold, request.Page2PointIsBulletPoint, legacyPointQueue);
        AddPointPage(data, 3, 11, request.Page3Points, request.Page3PointIsBold, request.Page3PointIsBulletPoint, legacyPointQueue);
        AddPointPage(data, 4, 45, request.Page4Points, request.Page4PointIsBold, null, legacyPointQueue);
        AddPointPage(data, 5, 35, request.Page5Points, request.Page5PointIsBold, null, legacyPointQueue);
        AddPointPage(data, 6, 11, request.Page6Points, request.Page6PointIsBold, request.Page6PointIsBulletPoint, legacyPointQueue);
        AddPointPage(data, 7, 11, request.Page7Points, request.Page7PointIsBold, request.Page7PointIsBulletPoint, legacyPointQueue);

        AddOrderedPlaceholders(
            data,
            MatrixNumericTokenOrder.Select(static number => $"{{{{num_{number}}}}}"),
            request.NumericValues,
            request.NumericValueIsBold);

        AddOrderedPlaceholders(
            data,
            Enumerable.Range(1, 8).Select(static number => $"{{{{perc_{number}%}}}}"),
            request.PercentageValues,
            request.PercentageValueIsBold);

        AddReplacement(
            data,
            "{{num_22}}",
            request.TotalStoriesCommitted,
            request.TotalStoriesCommittedIsBold,
            request.TotalStoriesCommittedIsBulletPoint);

        AddReplacement(
            data,
            "{{num_23}}",
            request.TotalStoryPoints,
            request.TotalStoryPointsIsBold,
            request.TotalStoryPointsIsBulletPoint);

        AddReplacement(
            data,
            "{{24}}",
            request.NewUserStories,
            request.NewUserStoriesIsBold,
            request.NewUserStoriesIsBulletPoint);

        AddReplacement(
            data,
            "{{25}}",
            request.SpilloverStories,
            request.SpilloverStoriesIsBold,
            request.SpilloverStoriesIsBulletPoint);

        AddReplacement(
            data,
            "{{26}}",
            request.AgileCeremonyProcessItem,
            request.AgileCeremonyProcessItemIsBold,
            request.AgileCeremonyProcessItemIsBulletPoint);

        return data;
    }

    private static void AddPointPage(
        IDictionary<string, PptTextValue> data,
        int pageNumber,
        int count,
        List<string>? values,
        IReadOnlyList<bool>? boldFlags,
        IReadOnlyList<bool>? bulletFlags,
        Queue<string> legacyPointQueue)
    {
        var resolvedValues = ResolvePointValues(values, boldFlags, bulletFlags, legacyPointQueue, count);

        for (var index = 1; index <= count; index++)
        {
            var token = $"{{{{POINT_{index}_PAGE_{pageNumber}}}}}";
            AddReplacement(data, token, resolvedValues[index - 1]);

            if (pageNumber == 7 && index == 9)
            {
                AddReplacement(data, "{{POINT_9_PAGE_7}", resolvedValues[index - 1]);
            }
        }
    }

    private static void AddOrderedPlaceholders(
        IDictionary<string, PptTextValue> data,
        IEnumerable<string> tokens,
        List<string>? values,
        IReadOnlyList<bool>? boldFlags)
    {
        var normalizedValues = values ?? [];
        var tokenList = tokens.ToList();

        for (var index = 0; index < tokenList.Count; index++)
        {
            var value = index < normalizedValues.Count ? normalizedValues[index] : string.Empty;
            AddReplacement(data, tokenList[index], value, GetFlag(boldFlags, index), null);
        }
    }

    private static List<PptTextValue> ResolvePointValues(
        List<string>? values,
        IReadOnlyList<bool>? boldFlags,
        IReadOnlyList<bool>? bulletFlags,
        Queue<string> legacyPointQueue,
        int count)
    {
        var result = new List<PptTextValue>(count);
        var normalizedValues = values ?? [];

        for (var index = 0; index < normalizedValues.Count && index < count; index++)
        {
            result.Add(new PptTextValue
            {
                Text = Normalize(normalizedValues[index]),
                IsBold = GetFlag(boldFlags, index),
                IsBulletPoint = bulletFlags is null ? null : GetFlag(bulletFlags, index, true)
            });
        }

        while (result.Count < count && legacyPointQueue.Count > 0)
        {
            result.Add(new PptTextValue
            {
                Text = legacyPointQueue.Dequeue(),
                IsBold = false,
                IsBulletPoint = bulletFlags is null ? null : true
            });
        }

        while (result.Count < count)
        {
            result.Add(new PptTextValue
            {
                Text = string.Empty,
                IsBold = false,
                IsBulletPoint = bulletFlags is null ? null : true
            });
        }

        return result;
    }

    private static async Task<Dictionary<string, string>> SaveImagesAsync(
        List<IFormFile>? files,
        ICollection<string> tempImagePaths)
    {
        var imageMappings = new Dictionary<string, string>(StringComparer.Ordinal);
        if (files is null || files.Count == 0)
        {
            return imageMappings;
        }

        var tokens = new[] { "{{img_1}}", "{{image_2}}", "{{image_3}}", "{{image_4}}" };

        for (var index = 0; index < files.Count && index < tokens.Length; index++)
        {
            var file = files[index];
            if (file is null || file.Length == 0)
            {
                continue;
            }

            var extension = Path.GetExtension(file.FileName);
            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = ".jpg";
            }

            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");

            await using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await file.CopyToAsync(stream);
            }

            tempImagePaths.Add(tempPath);
            imageMappings[tokens[index]] = tempPath;
        }

        return imageMappings;
    }

    private static void AddReplacement(
        IDictionary<string, PptTextValue> data,
        string token,
        PptTextValue value)
    {
        data[token] = value;
    }

    private static void AddReplacement(
        IDictionary<string, PptTextValue> data,
        string token,
        string? text,
        bool? isBold,
        bool? isBulletPoint = null)
    {
        data[token] = new PptTextValue
        {
            Text = Normalize(text),
            IsBold = isBold,
            IsBulletPoint = isBulletPoint
        };
    }

    private static bool GetFlag(IReadOnlyList<bool>? values, int index, bool defaultValue = false)
    {
        return values is not null && index < values.Count ? values[index] : defaultValue;
    }

    private static string Normalize(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
    }

    private static string BuildDownloadFileName(string? date)
    {
        var normalizedDate = Normalize(date);
        if (string.IsNullOrWhiteSpace(normalizedDate))
        {
            return "Pegasus_sprint_retro.pptx";
        }

        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitizedDate = new string(
            normalizedDate
                .Select(character => invalidCharacters.Contains(character) ? '_' : character)
                .ToArray());

        sanitizedDate = string.Join(
            "_",
            sanitizedDate
                .Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries));

        return $"Pegasus_sprint_retro{sanitizedDate}.pptx";
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
        {
            return;
        }

        try
        {
            System.IO.File.Delete(path);
        }
        catch
        {
            // Cleanup should never fail the PPT generation response.
        }
    }
}
