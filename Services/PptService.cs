using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

// ✅ Aliases to avoid ambiguity
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ppt_mapper.Services
{
    public class PptService
    {
        // =========================
        // TEXT REPLACEMENT
        // =========================
        public void ReplacePlaceholders(string filePath, Dictionary<string, string> data)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, true))
            {
                var slides = ppt.PresentationPart.SlideParts;

                foreach (var slide in slides)
                {
                    var paragraphs = slide.Slide.Descendants<A.Paragraph>();

                    foreach (var para in paragraphs)
                    {
                        var fullText = string.Join("", para.Descendants<A.Text>().Select(t => t.Text));

                        foreach (var item in data)
                        {
                            if (fullText.Contains(item.Key))
                            {
                                fullText = fullText.Replace(item.Key, item.Value);

                                var texts = para.Descendants<A.Text>().ToList();
                                if (texts.Count > 0)
                                {
                                    texts.First().Text = fullText;

                                    for (int i = 1; i < texts.Count; i++)
                                        texts[i].Text = "";
                                }
                            }
                        }
                    }
                }
            }
        }
        public void ReplaceImages(string filePath, List<string> imagePaths)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, true))
            {
                var slides = ppt.PresentationPart.SlideParts.ToList();
                int index = 0;

                foreach (var slide in slides)
                {
                    var shapes = slide.Slide.Descendants<P.Shape>().ToList();

                    foreach (var shape in shapes)
                    {
                        var textElement = shape.TextBody?.Descendants<A.Text>().FirstOrDefault();
                        if (textElement == null) continue;

                        if (!textElement.Text.Contains("{{image")) continue;
                        if (index >= imagePaths.Count) continue;

                        var imagePath = imagePaths[index];
                        if (!File.Exists(imagePath)) continue;
                        if (new FileInfo(imagePath).Length == 0) continue;

                        // 🔥 Get placeholder position
                        var transform = shape.ShapeProperties?.Transform2D;
                        var offset = transform?.Offset;
                        var extents = transform?.Extents;

                        long x = offset?.X ?? 0;
                        long y = offset?.Y ?? 0;
                        long cx = extents?.Cx ?? 3000000;
                        long cy = extents?.Cy ?? 3000000;

                        // 🔥 Detect image type
                        var imagePart = slide.AddImagePart(GetImagePartType(imagePath));

                        using (var stream = File.OpenRead(imagePath))
                        {
                            imagePart.FeedData(stream);
                        }

                        var relId = slide.GetIdOfPart(imagePart);

                        // 🔥 Generate UNIQUE ID
                        var maxId = slide.Slide.Descendants<P.NonVisualDrawingProperties>()
                            .Max(x => (uint?)x.Id.Value) ?? 0;

                        var newId = maxId + 1;

                        // 🔥 Create image at SAME position
                        var picture = new P.Picture(
                            new P.NonVisualPictureProperties(
                                new P.NonVisualDrawingProperties()
                                {
                                    Id = new UInt32Value(newId),
                                    Name = "Image" + newId
                                },
                                new P.NonVisualPictureDrawingProperties()
                            ),
                            new P.BlipFill(
                                new A.Blip() { Embed = relId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new P.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = x, Y = y },
                                    new A.Extents() { Cx = cx, Cy = cy }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        );

                        // 🔥 Replace placeholder safely
                        var parent = shape.Parent;
                        parent.InsertAfter(picture, shape);
                        parent.RemoveChild(shape);

                        index++;
                    }
                }
            }
        }

        // =========================
        // IMAGE TYPE DETECTION
        // =========================
        private ImagePartType GetImagePartType(string path)
        {
            var ext = Path.GetExtension(path).ToLower();

            return ext switch
            {
                ".png" => ImagePartType.Png,
                ".jpg" => ImagePartType.Jpeg,
                ".jpeg" => ImagePartType.Jpeg,
                _ => ImagePartType.Jpeg
            };
        }
    }
}