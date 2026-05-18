using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ppt_mapper.Models;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ppt_mapper.Services
{
    public class PptService
    {
        private static readonly string[] ThreeImageTokens = ["{{image_2}}", "{{image_3}}", "{{image_4}}"];

        private const long DefaultSlideWidth = 12192000;
        private const long DefaultSlideHeight = 6858000;
        private const long FooterTop = 6628130;
        private const int DefaultBulletMargin = 285750;
        private const int DefaultBulletIndent = -285750;

        public void ReplacePlaceholders(string filePath, IReadOnlyDictionary<string, PptTextValue> data)
        {
            using var ppt = PresentationDocument.Open(filePath, true);
            var slides = ppt.PresentationPart?.SlideParts ?? [];

            foreach (var slide in slides)
            {
                foreach (var paragraph in slide.Slide.Descendants<A.Paragraph>().ToList())
                {
                    ReplaceParagraph(paragraph, data);
                }
            }
        }

        public void ReplaceImages(string filePath, IReadOnlyDictionary<string, string> imagePaths)
        {
            using var ppt = PresentationDocument.Open(filePath, true);
            var presentationPart = ppt.PresentationPart;
            if (presentationPart?.Presentation is null)
            {
                return;
            }

            var slideWidth = presentationPart.Presentation.SlideSize?.Cx?.Value ?? DefaultSlideWidth;
            var slideHeight = presentationPart.Presentation.SlideSize?.Cy?.Value ?? DefaultSlideHeight;

            foreach (var slidePart in presentationPart.SlideParts)
            {
                var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

                ReplaceSingleImage(slidePart, shapes, imagePaths, slideWidth, slideHeight);
                ReplaceThreeImageLayout(slidePart, shapes, imagePaths, slideWidth, slideHeight);
            }
        }

        private static void ReplaceParagraph(A.Paragraph paragraph, IReadOnlyDictionary<string, PptTextValue> replacements)
        {
            var runInfos = GetRunInfos(paragraph);
            if (runInfos.Count == 0)
            {
                return;
            }

            var currentText = string.Concat(runInfos.Select(static run => run.Text));
            if (string.IsNullOrEmpty(currentText))
            {
                return;
            }

            var matches = FindMatches(currentText, replacements);
            if (matches.Count == 0)
            {
                return;
            }

            var bulletSetting = GetBulletSetting(matches, replacements);
            if (ShouldRemoveParagraph(matches, replacements, bulletSetting))
            {
                paragraph.Remove();
                return;
            }

            var newParagraph = new A.Paragraph();
            if (paragraph.ParagraphProperties is not null)
            {
                newParagraph.Append((A.ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true));
            }

            ApplyBulletFormatting(newParagraph, bulletSetting);

            var cursor = 0;
            foreach (var match in matches)
            {
                if (match.Index > cursor)
                {
                    AppendStaticRuns(newParagraph, runInfos, cursor, match.Index);
                }

                AppendReplacementRun(
                    newParagraph,
                    runInfos,
                    match.Index,
                    match.Index + match.Token.Length,
                    replacements[match.Token]);

                cursor = match.Index + match.Token.Length;
            }

            if (cursor < currentText.Length)
            {
                AppendStaticRuns(newParagraph, runInfos, cursor, currentText.Length);
            }

            var endParagraphProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (newParagraph.GetFirstChild<A.EndParagraphRunProperties>() is null && endParagraphProperties is not null)
            {
                newParagraph.Append((A.EndParagraphRunProperties)endParagraphProperties.CloneNode(true));
            }

            paragraph.Parent?.ReplaceChild(newParagraph, paragraph);
        }

        private static List<RunInfo> GetRunInfos(A.Paragraph paragraph)
        {
            var result = new List<RunInfo>();
            var cursor = 0;

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var text = string.Concat(run.Elements<A.Text>().Select(static value => value.Text));
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                result.Add(new RunInfo(
                    text,
                    cursor,
                    cursor + text.Length,
                    run.RunProperties is null ? null : (A.RunProperties)run.RunProperties.CloneNode(true)));

                cursor += text.Length;
            }

            return result;
        }

        private static List<ParagraphMatch> FindMatches(string text, IReadOnlyDictionary<string, PptTextValue> replacements)
        {
            var matches = new List<ParagraphMatch>();
            var tokens = replacements.Keys.ToList();
            var cursor = 0;

            while (cursor < text.Length)
            {
                var nextIndex = -1;
                string? nextToken = null;

                foreach (var token in tokens)
                {
                    var index = text.IndexOf(token, cursor, StringComparison.Ordinal);
                    if (index < 0)
                    {
                        continue;
                    }

                    if (nextIndex == -1 ||
                        index < nextIndex ||
                        (index == nextIndex && token.Length > (nextToken?.Length ?? 0)))
                    {
                        nextIndex = index;
                        nextToken = token;
                    }
                }

                if (nextIndex < 0 || nextToken is null)
                {
                    break;
                }

                matches.Add(new ParagraphMatch(nextToken, nextIndex));
                cursor = nextIndex + nextToken.Length;
            }

            return matches;
        }

        private static void AppendStaticRuns(A.Paragraph paragraph, IReadOnlyList<RunInfo> runInfos, int start, int end)
        {
            foreach (var runInfo in runInfos)
            {
                var segmentStart = Math.Max(start, runInfo.Start);
                var segmentEnd = Math.Min(end, runInfo.End);
                if (segmentStart >= segmentEnd)
                {
                    continue;
                }

                var text = runInfo.Text.Substring(segmentStart - runInfo.Start, segmentEnd - segmentStart);
                paragraph.Append(CreateRun(text, runInfo.RunProperties));
            }
        }

        private static void AppendReplacementRun(
            A.Paragraph paragraph,
            IReadOnlyList<RunInfo> runInfos,
            int start,
            int end,
            PptTextValue replacement)
        {
            if (string.IsNullOrEmpty(replacement.Text))
            {
                return;
            }

            var baseRunInfo = runInfos.FirstOrDefault(static run => run.Start < run.End && run.End > run.Start);
            foreach (var runInfo in runInfos)
            {
                if (runInfo.End > start && runInfo.Start < end)
                {
                    baseRunInfo = runInfo;
                    break;
                }
            }

            var runProperties = baseRunInfo.RunProperties is null
                ? new A.RunProperties()
                : (A.RunProperties)baseRunInfo.RunProperties.CloneNode(true);

            if (replacement.IsBold.HasValue)
            {
                runProperties.Bold = replacement.IsBold.Value;
            }

            paragraph.Append(CreateRun(replacement.Text, runProperties));
        }

        private static A.Run CreateRun(string text, A.RunProperties? runProperties)
        {
            var run = new A.Run();
            if (runProperties is not null)
            {
                run.Append((A.RunProperties)runProperties.CloneNode(true));
            }

            run.Append(new A.Text(text));
            return run;
        }

        private static bool ShouldRemoveParagraph(
            IReadOnlyList<ParagraphMatch> matches,
            IReadOnlyDictionary<string, PptTextValue> replacements,
            bool? bulletSetting)
        {
            return bulletSetting == true &&
                   matches.All(match => string.IsNullOrWhiteSpace(replacements[match.Token].Text));
        }

        private static bool? GetBulletSetting(
            IReadOnlyList<ParagraphMatch> matches,
            IReadOnlyDictionary<string, PptTextValue> replacements)
        {
            foreach (var match in matches)
            {
                var value = replacements[match.Token];
                if (value.IsBulletPoint.HasValue)
                {
                    return value.IsBulletPoint.Value;
                }
            }

            return null;
        }

        private static void ApplyBulletFormatting(
            A.Paragraph paragraph,
            bool? bulletSetting)
        {
            if (!bulletSetting.HasValue)
            {
                return;
            }

            var paragraphProperties = paragraph.GetFirstChild<A.ParagraphProperties>();
            if (paragraphProperties is null)
            {
                paragraphProperties = new A.ParagraphProperties();
                paragraph.PrependChild(paragraphProperties);
            }

            if (bulletSetting.Value)
            {
                paragraphProperties.RemoveAllChildren<A.NoBullet>();

                if (paragraphProperties.GetFirstChild<A.CharacterBullet>() is null &&
                    paragraphProperties.GetFirstChild<A.AutoNumberedBullet>() is null)
                {
                    paragraphProperties.Append(new A.BulletFont { Typeface = "Arial" });
                    paragraphProperties.Append(new A.CharacterBullet { Char = "•" });
                }

                paragraphProperties.LeftMargin ??= DefaultBulletMargin;
                paragraphProperties.Indent ??= DefaultBulletIndent;
            }
            else
            {
                paragraphProperties.RemoveAllChildren<A.CharacterBullet>();
                paragraphProperties.RemoveAllChildren<A.AutoNumberedBullet>();
                paragraphProperties.RemoveAllChildren<A.BulletFont>();
                paragraphProperties.RemoveAllChildren<A.NoBullet>();
                paragraphProperties.Append(new A.NoBullet());
                paragraphProperties.LeftMargin = 0;
                paragraphProperties.Indent = 0;
            }
        }

        private static void ReplaceSingleImage(
            SlidePart slidePart,
            IReadOnlyList<P.Shape> shapes,
            IReadOnlyDictionary<string, string> imagePaths,
            long slideWidth,
            long slideHeight)
        {
            var placeholder = shapes
                .Select(shape => new { Shape = shape, Token = GetShapeText(shape) })
                .FirstOrDefault(static item => item.Token == "{{img_1}}");

            if (placeholder is null)
            {
                return;
            }

            if (!imagePaths.TryGetValue("{{img_1}}", out var imagePath) ||
                string.IsNullOrWhiteSpace(imagePath) ||
                !File.Exists(imagePath) ||
                new FileInfo(imagePath).Length == 0)
            {
                RemoveShape(placeholder.Shape);
                return;
            }

            var titleBottom = GetTitleBottom(slidePart);
            var bounds = new EmuRectangle(
                900000,
                titleBottom + 450000,
                slideWidth - 1800000,
                Math.Max(1000000, Math.Min(FooterTop, slideHeight - 200000) - (titleBottom + 600000)));

            ReplaceShapeWithImage(slidePart, placeholder.Shape, imagePath, bounds);
        }

        private static void ReplaceThreeImageLayout(
            SlidePart slidePart,
            IReadOnlyList<P.Shape> shapes,
            IReadOnlyDictionary<string, string> imagePaths,
            long slideWidth,
            long slideHeight)
        {
            var placeholders = shapes
                .Select(shape => new ImagePlaceholder(shape, GetShapeText(shape)))
                .Where(static item => ThreeImageTokens.Contains(item.Token, StringComparer.Ordinal))
                .OrderBy(static item => Array.IndexOf(ThreeImageTokens, item.Token))
                .ToList();

            if (placeholders.Count == 0)
            {
                return;
            }

            var validImages = placeholders
                .Select(static item => item.Token)
                .Where(token => imagePaths.TryGetValue(token, out var path) &&
                                !string.IsNullOrWhiteSpace(path) &&
                                File.Exists(path) &&
                                new FileInfo(path).Length > 0)
                .Select(token => imagePaths[token])
                .Take(3)
                .ToList();

            var parent = placeholders[0].Shape.Parent;

            foreach (var placeholder in placeholders)
            {
                RemoveShape(placeholder.Shape);
            }

            if (validImages.Count == 0)
            {
                return;
            }

            if (parent is null)
            {
                return;
            }

            var titleBottom = GetTitleBottom(slidePart);
            var bounds = ResolveThreeImageBounds(validImages.Count, titleBottom, slideWidth, slideHeight);

            for (var index = 0; index < validImages.Count && index < bounds.Count; index++)
            {
                var picture = CreatePicture(slidePart, validImages[index], bounds[index]);
                parent.AppendChild(picture);
            }
        }

        private static IReadOnlyList<EmuRectangle> ResolveThreeImageBounds(
            int count,
            long titleBottom,
            long slideWidth,
            long slideHeight)
        {
            var contentTop = titleBottom + 650000;
            var contentBottom = Math.Min(FooterTop, slideHeight - 200000);
            var contentHeight = Math.Max(1000000, contentBottom - contentTop - 150000);

            if (count <= 1)
            {
                var width = Math.Min(slideWidth - 2400000, 7200000);
                var x = (slideWidth - width) / 2;
                return [new EmuRectangle(x, contentTop, width, contentHeight)];
            }

            if (count == 2)
            {
                const long outerPadding = 700000;
                const long gap = 400000;
                var width = (slideWidth - (outerPadding * 2) - gap) / 2;

                return
                [
                    new EmuRectangle(outerPadding, contentTop, width, contentHeight),
                    new EmuRectangle(outerPadding + width + gap, contentTop, width, contentHeight)
                ];
            }

            const long leftPadding = 450000;
            const long rightPadding = 450000;
            const long threeGap = 250000;
            var slotWidth = (slideWidth - leftPadding - rightPadding - (threeGap * 2)) / 3;

            return
            [
                new EmuRectangle(leftPadding, contentTop, slotWidth, contentHeight),
                new EmuRectangle(leftPadding + slotWidth + threeGap, contentTop, slotWidth, contentHeight),
                new EmuRectangle(leftPadding + ((slotWidth + threeGap) * 2), contentTop, slotWidth, contentHeight)
            ];
        }

        private static void ReplaceShapeWithImage(
            SlidePart slidePart,
            P.Shape shape,
            string imagePath,
            EmuRectangle bounds)
        {
            var picture = CreatePicture(slidePart, imagePath, bounds);
            var parent = shape.Parent;
            if (parent is null)
            {
                return;
            }

            parent.InsertAfter(picture, shape);
            parent.RemoveChild(shape);
        }

        private static P.Picture CreatePicture(SlidePart slidePart, string imagePath, EmuRectangle bounds)
        {
            var imagePart = slidePart.AddImagePart(GetImagePartType(imagePath));
            using (var stream = File.OpenRead(imagePath))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var pictureId = (slidePart.Slide.Descendants<P.NonVisualDrawingProperties>()
                .Max(static value => (uint?)value.Id?.Value) ?? 0) + 1;

            return new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties
                    {
                        Id = pictureId,
                        Name = $"MappedImage{pictureId}"
                    },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new P.ApplicationNonVisualDrawingProperties()),
                BuildBlipFill(relationshipId, imagePath, bounds),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = bounds.X, Y = bounds.Y },
                        new A.Extents { Cx = bounds.Width, Cy = bounds.Height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        }

        private static P.BlipFill BuildBlipFill(string relationshipId, string imagePath, EmuRectangle targetBounds)
        {
            var blipFill = new P.BlipFill
            {
                Blip = new A.Blip { Embed = relationshipId }
            };

            var sourceRectangle = BuildSourceRectangle(imagePath, targetBounds);
            if (sourceRectangle is not null)
            {
                blipFill.SourceRectangle = sourceRectangle;
            }

            blipFill.Append(new A.Stretch(new A.FillRectangle()));
            return blipFill;
        }

        private static A.SourceRectangle? BuildSourceRectangle(string imagePath, EmuRectangle targetBounds)
        {
            if (!TryGetImageSize(imagePath, out var imageWidth, out var imageHeight) ||
                imageWidth <= 0 ||
                imageHeight <= 0 ||
                targetBounds.Width <= 0 ||
                targetBounds.Height <= 0)
            {
                return null;
            }

            var imageAspectRatio = (double)imageWidth / imageHeight;
            var targetAspectRatio = (double)targetBounds.Width / targetBounds.Height;

            var left = 0;
            var right = 0;
            var top = 0;
            var bottom = 0;

            if (imageAspectRatio > targetAspectRatio)
            {
                var keptWidthRatio = targetAspectRatio / imageAspectRatio;
                var cropRatio = (1d - keptWidthRatio) / 2d;
                left = ToCropValue(cropRatio);
                right = ToCropValue(cropRatio);
            }
            else if (imageAspectRatio < targetAspectRatio)
            {
                var keptHeightRatio = imageAspectRatio / targetAspectRatio;
                var cropRatio = (1d - keptHeightRatio) / 2d;
                top = ToCropValue(cropRatio);
                bottom = ToCropValue(cropRatio);
            }

            if (left == 0 && right == 0 && top == 0 && bottom == 0)
            {
                return null;
            }

            return new A.SourceRectangle
            {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom
            };
        }

        private static int ToCropValue(double ratio)
        {
            return (int)Math.Clamp(Math.Round(ratio * 100000d), 0d, 50000d);
        }

        private static bool TryGetImageSize(string imagePath, out int width, out int height)
        {
            width = 0;
            height = 0;

            try
            {
                using var stream = File.OpenRead(imagePath);
                using var reader = new BinaryReader(stream);

                if (TryReadPngSize(reader, out width, out height) ||
                    TryReadGifSize(reader, out width, out height) ||
                    TryReadBmpSize(reader, out width, out height) ||
                    TryReadJpegSize(reader, out width, out height))
                {
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static long GetTitleBottom(SlidePart slidePart)
        {
            long titleBottom = 0;

            foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
            {
                var text = GetShapeText(shape);
                if (!text.Contains("Sprint", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var bounds = GetBounds(shape);
                titleBottom = Math.Max(titleBottom, bounds.Bottom);
            }

            return titleBottom == 0 ? 1250000 : titleBottom;
        }

        private static string GetShapeText(P.Shape shape)
        {
            return string.Concat(shape.TextBody?.Descendants<A.Text>().Select(static text => text.Text) ?? []);
        }

        private static EmuRectangle GetBounds(P.Shape shape)
        {
            var transform = shape.ShapeProperties?.Transform2D;
            var offset = transform?.Offset;
            var extents = transform?.Extents;

            return new EmuRectangle(
                offset?.X?.Value ?? 0,
                offset?.Y?.Value ?? 0,
                extents?.Cx?.Value ?? 0,
                extents?.Cy?.Value ?? 0);
        }

        private static void RemoveShape(P.Shape shape)
        {
            var parent = shape.Parent;
            if (parent is not null)
            {
                parent.RemoveChild(shape);
            }
        }

        private static ImagePartType GetImagePartType(string path)
        {
            return Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".png" => ImagePartType.Png,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                ".tif" => ImagePartType.Tiff,
                ".tiff" => ImagePartType.Tiff,
                ".jpg" => ImagePartType.Jpeg,
                ".jpeg" => ImagePartType.Jpeg,
                _ => ImagePartType.Jpeg
            };
        }

        private readonly record struct EmuRectangle(long X, long Y, long Width, long Height)
        {
            public long Bottom => Y + Height;
        }

        private readonly record struct RunInfo(string Text, int Start, int End, A.RunProperties? RunProperties);

        private readonly record struct ParagraphMatch(string Token, int Index);

        private readonly record struct ImagePlaceholder(P.Shape Shape, string Token);

        private static bool TryReadPngSize(BinaryReader reader, out int width, out int height)
        {
            width = 0;
            height = 0;

            reader.BaseStream.Position = 0;
            var signature = reader.ReadBytes(8);
            if (signature.Length < 8 ||
                signature[0] != 137 ||
                signature[1] != 80 ||
                signature[2] != 78 ||
                signature[3] != 71)
            {
                return false;
            }

            reader.BaseStream.Position = 16;
            width = ReadInt32BigEndian(reader);
            height = ReadInt32BigEndian(reader);
            return width > 0 && height > 0;
        }

        private static bool TryReadGifSize(BinaryReader reader, out int width, out int height)
        {
            width = 0;
            height = 0;

            reader.BaseStream.Position = 0;
            var header = reader.ReadBytes(6);
            if (header.Length < 6)
            {
                return false;
            }

            var signature = System.Text.Encoding.ASCII.GetString(header);
            if (!string.Equals(signature, "GIF87a", StringComparison.Ordinal) &&
                !string.Equals(signature, "GIF89a", StringComparison.Ordinal))
            {
                return false;
            }

            width = reader.ReadUInt16();
            height = reader.ReadUInt16();
            return width > 0 && height > 0;
        }

        private static bool TryReadBmpSize(BinaryReader reader, out int width, out int height)
        {
            width = 0;
            height = 0;

            reader.BaseStream.Position = 0;
            if (reader.ReadByte() != 'B' || reader.ReadByte() != 'M')
            {
                return false;
            }

            reader.BaseStream.Position = 18;
            width = reader.ReadInt32();
            height = Math.Abs(reader.ReadInt32());
            return width > 0 && height > 0;
        }

        private static bool TryReadJpegSize(BinaryReader reader, out int width, out int height)
        {
            width = 0;
            height = 0;

            reader.BaseStream.Position = 0;
            if (reader.ReadByte() != 0xFF || reader.ReadByte() != 0xD8)
            {
                return false;
            }

            while (reader.BaseStream.Position < reader.BaseStream.Length)
            {
                if (reader.ReadByte() != 0xFF)
                {
                    return false;
                }

                byte marker;
                do
                {
                    marker = reader.ReadByte();
                } while (marker == 0xFF);

                if (marker == 0xD9 || marker == 0xDA)
                {
                    return false;
                }

                var segmentLength = ReadUInt16BigEndian(reader);
                if (segmentLength < 2)
                {
                    return false;
                }

                if (marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF)
                {
                    _ = reader.ReadByte();
                    height = ReadUInt16BigEndian(reader);
                    width = ReadUInt16BigEndian(reader);
                    return width > 0 && height > 0;
                }

                reader.BaseStream.Seek(segmentLength - 2, SeekOrigin.Current);
            }

            return false;
        }

        private static int ReadInt32BigEndian(BinaryReader reader)
        {
            var bytes = reader.ReadBytes(sizeof(int));
            if (bytes.Length < sizeof(int))
            {
                return 0;
            }

            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(bytes);
            }

            return BitConverter.ToInt32(bytes, 0);
        }

        private static ushort ReadUInt16BigEndian(BinaryReader reader)
        {
            var bytes = reader.ReadBytes(sizeof(ushort));
            if (bytes.Length < sizeof(ushort))
            {
                return 0;
            }

            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(bytes);
            }

            return BitConverter.ToUInt16(bytes, 0);
        }
    }
}
