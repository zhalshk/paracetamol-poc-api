using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using iText.Layout.Element;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfiumViewer;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Tesseract;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Color = System.Drawing.Color;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using PdfiumDocument = PdfiumViewer.PdfDocument;
using PdfpigDocument = UglyToad.PdfPig.PdfDocument;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace paracetamol_poc_api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OcrController : ControllerBase
    {

        private readonly string tessDataPath = Path.Combine(Directory.GetCurrentDirectory(), "tessdata");

        [HttpPost("extract")]
        public IActionResult Extract([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            using (var stream = file.OpenReadStream())
            using (var pdfDoc = PdfiumDocument.Load(stream))
            using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
            {
                for (int i = 0; i < pdfDoc.PageCount; i++)
                {
                    using (var pageImg = pdfDoc.Render(i, 500, 500, true)) // high DPI for small text
                    {
                        // Convert page to base64
                        string pageBase64;
                        using (var msImg = new MemoryStream())
                        {
                            pageImg.Save(msImg, ImageFormat.Png);
                            pageBase64 = Convert.ToBase64String(msImg.ToArray());
                        }

                        // OCR for text
                        string text;
                        using (var msPix = new MemoryStream())
                        {
                            pageImg.Save(msPix, ImageFormat.Png);
                            msPix.Position = 0;
                            using (var pix = Pix.LoadFromMemory(msPix.ToArray()))
                            using (var page = engine.Process(pix))
                            {
                                text = page.GetText() ?? "";
                                text = Regex.Replace(text, @"\s*\n\s*", "\n").Trim();
                            }
                        }

                        resultPages.Add(new
                        {
                            Page = i + 1,
                            Text = text,
                            PageImageBase64 = pageBase64
                        });
                    }
                }
            }

            return Ok(new { Pages = resultPages });
        }

        [HttpPost("extract-handwriting")]
        public IActionResult ExtractHandwriting([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            using var stream = file.OpenReadStream();
            using var pdfDoc = PdfiumViewer.PdfDocument.Load(stream);

            // Load Tesseract with best-trained data
            using var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.LstmOnly);
            engine.SetVariable("user_defined_dpi", "300");
            engine.SetVariable("preserve_interword_spaces", "1");
            engine.DefaultPageSegMode = PageSegMode.Auto;

            for (int i = 0; i < pdfDoc.PageCount; i++)
            {
                using var pageImg = pdfDoc.Render(i, 500, 500, true); // 500 DPI for handwriting
                using var processedImg = PreprocessForOCR((Bitmap)pageImg);

                string text = "";

                // Convert processed image into Pix
                using (var msPix = new MemoryStream())
                {
                    processedImg.Save(msPix, ImageFormat.Png);
                    msPix.Position = 0;

                    using var pix = Pix.LoadFromMemory(msPix.ToArray());

                    // Try multiple segmentation modes for better accuracy
                    using (var page = engine.Process(pix, PageSegMode.SparseText))
                    {
                        text = page.GetText() ?? "";
                    }

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                        {
                            text = page.GetText() ?? "";
                        }
                    }

                    text = Regex.Replace(text, @"\s*\n\s*", "\n").Trim();
                }

                resultPages.Add(new
                {
                    Page = i + 1,
                    Text = text
                });
            }

            return Ok(new { Pages = resultPages });
        }

        // ✅ Image preprocessing for handwriting OCR
        private static Bitmap PreprocessForOCR(Bitmap input)
        {
            Bitmap gray = new Bitmap(input.Width, input.Height, PixelFormat.Format24bppRgb);

            // Convert to grayscale
            using (Graphics g = Graphics.FromImage(gray))
            {
                var colorMatrix = new System.Drawing.Imaging.ColorMatrix(
                    new float[][]
                    {
                        new float[] {0.3f, 0.3f, 0.3f, 0, 0},
                        new float[] {0.59f, 0.59f, 0.59f, 0, 0},
                        new float[] {0.11f, 0.11f, 0.11f, 0, 0},
                        new float[] {0, 0, 0, 1, 0},
                        new float[] {0, 0, 0, 0, 1}
                    });
                var attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrix);

                g.DrawImage(input, new Rectangle(0, 0, input.Width, input.Height),
                    0, 0, input.Width, input.Height, GraphicsUnit.Pixel, attributes);
            }

            // Simple binarization (thresholding)
            for (int y = 0; y < gray.Height; y++)
            {
                for (int x = 0; x < gray.Width; x++)
                {
                    var pixel = gray.GetPixel(x, y);
                    int brightness = (pixel.R + pixel.G + pixel.B) / 3;
                    Color newColor = brightness > 140 ? Color.White : Color.Black;
                    gray.SetPixel(x, y, newColor);
                }
            }

            return gray;
        }


        [HttpPost("extract-pdfpig")]
        public IActionResult ExtractPdfPig([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            // Copy uploaded file to a byte array once
            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                file.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            // PdfPig for text
            using (var msPdfPig = new MemoryStream(fileBytes))
            using (var pdf = PdfpigDocument.Open(msPdfPig))
            {
                // Pdfium for images
                using (var msPdfium = new MemoryStream(fileBytes))
                using (var pdfiumDoc = PdfiumViewer.PdfDocument.Load(msPdfium))
                {
                    for (int i = 0; i < pdf.NumberOfPages; i++)
                    {
                        var page = pdf.GetPage(i + 1);
                        string text = page.Text;

                        var pageImg = pdfiumDoc.Render(i, 400, 400, true);
                        string pageImageBase64;
                        using (var msImg = new MemoryStream())
                        {
                            pageImg.Save(msImg, System.Drawing.Imaging.ImageFormat.Png);
                            pageImageBase64 = Convert.ToBase64String(msImg.ToArray());
                        }

                        resultPages.Add(new
                        {
                            Page = i + 1,
                            Text = text,
                            PageImageBase64 = pageImageBase64
                        });
                    }
                }
            }

            return Ok(new { Pages = resultPages });
        }



        [HttpPost("extract-pdfpig-tables")]
        public IActionResult ExtractPdfPigTables([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            using (var ms = new MemoryStream())
            {
                file.CopyTo(ms);
                ms.Position = 0;

                using var pdf = PdfpigDocument.Open(ms);

                for (int i = 0; i < pdf.NumberOfPages; i++)
                {
                    var page = pdf.GetPage(i + 1);

                    // Extract all words with coordinates
                    var words = page.GetWords().Select(w => new
                    {
                        Text = w.Text,
                        X = w.BoundingBox.Left,
                        Y = w.BoundingBox.Top
                    }).ToList();

                    // Group words by approximate Y coordinate to form rows
                    var rowGroups = words
                        .GroupBy(w => Math.Round(w.Y / 5.0) * 5) // group by 5 units tolerance
                        .OrderByDescending(g => g.Key) // top to bottom
                        .ToList();

                    var tableRows = new List<List<string>>();
                    foreach (var row in rowGroups)
                    {
                        // sort words left to right
                        var orderedRow = row.OrderBy(w => w.X).Select(w => w.Text).ToList();
                        tableRows.Add(orderedRow);
                    }

                    // Extract images
                    var images = new List<object>();
                    foreach (var img in page.GetImages())
                    {
                        images.Add(new
                        {
                            ImageBase64 = Convert.ToBase64String(img.RawBytes),
                            Width = img.WidthInSamples,
                            Height = img.HeightInSamples
                        });
                    }

                    resultPages.Add(new
                    {
                        Page = i + 1,
                        Words = words,
                        TableRows = tableRows,
                        Images = images
                    });
                }
            }

            return Ok(new { Pages = resultPages });
        }

        [HttpPost("extract-itext")]
        public IActionResult ExtractIText([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            using var ms = new MemoryStream();
            file.CopyTo(ms);
            ms.Position = 0;

            using var pdfReader = new PdfReader(ms);
            using var pdfDoc = new iText.Kernel.Pdf.PdfDocument(pdfReader);

            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                var page = pdfDoc.GetPage(i);

                // Extract text
                var strategy = new SimpleTextExtractionStrategy();
                string text = PdfTextExtractor.GetTextFromPage(page, strategy);

                // Extract images
                var listener = new ImageExtractionListener();
                var parser = new PdfCanvasProcessor(listener);
                parser.ProcessPageContent(page);

                var images = listener.ExtractedImages
                    .Select(img => new { img.Base64, img.Width, img.Height })
                    .ToList();

                resultPages.Add(new
                {
                    Page = i,
                    Text = text,
                    Images = images
                });
            }

            return Ok(new { Pages = resultPages });
        }

        [HttpPost("extract-itext-detailed")]
        public IActionResult ExtractITextDetailed([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            var resultPages = new List<object>();

            using var ms = new MemoryStream();
            file.CopyTo(ms);
            ms.Position = 0;

            using var pdfReader = new PdfReader(ms);
            using var pdfDoc = new iText.Kernel.Pdf.PdfDocument(pdfReader);

            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                var page = pdfDoc.GetPage(i);

                // Text + positions
                var strategy = new LocationTextExtractionStrategyEx();
                var parser = new PdfCanvasProcessor(strategy);
                parser.ProcessPageContent(page);

                var textChunks = strategy.GetChunks()
                    .Select(c => new
                    {
                        c.Text,
                        c.FontSize,
                        c.FontName,
                        c.X,
                        c.Y,
                        c.Width,
                        c.Height,
                        c.IsBold,
                        c.IsItalic
                    })
                    .ToList();

                // Extract images like before
                var imageListener = new ImageExtractionListener();
                var imageParser = new PdfCanvasProcessor(imageListener);
                imageParser.ProcessPageContent(page);

                var images = imageListener.ExtractedImages
                    .Select(img => new { img.Base64, img.Width, img.Height })
                    .ToList();

                resultPages.Add(new
                {
                    Page = i,
                    TextChunks = textChunks,
                    Images = images
                });
            }

            return Ok(new { Pages = resultPages });
        }

        [HttpPost("rebuild-docx")]
        public IActionResult RebuildDocx([FromBody] JsonElement json)
        {
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), $"Rebuilt_{Guid.NewGuid()}.docx");


            using (var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                var body = mainPart.Document.Body;

                foreach (var page in json.GetProperty("pages").EnumerateArray())
                {
                    foreach (var chunk in page.GetProperty("textChunks").EnumerateArray())
                    {
                        string text = chunk.GetProperty("text").GetString();
                        float fontSize = (float)chunk.GetProperty("fontSize").GetDouble();
                        bool isBold = chunk.GetProperty("isBold").GetBoolean();
                        bool isItalic = chunk.GetProperty("isItalic").GetBoolean();

                        var runProps = new RunProperties(
                            new FontSize() { Val = (fontSize * 2).ToString() } // Word uses half-points
                        );

                        if (isBold) runProps.Append(new Bold());
                        if (isItalic) runProps.Append(new Italic());

                        var run = new Run(runProps, new Text(text));
                        var para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

                        body.Append(para);
                    }

                    // Add images
                    foreach (var img in page.GetProperty("images").EnumerateArray())
                    {
                        var imgBytes = Convert.FromBase64String(img.GetProperty("base64").GetString());
                        AddImageToWord(mainPart, body, imgBytes);
                    }

                    // Page break
                    body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }
            }

            var fileBytes = System.IO.File.ReadAllBytes(outputPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Rebuilt.docx");
        }

        private void AddImageToWord(MainDocumentPart mainPart, Body body, byte[] imageBytes)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);

            var element =
                 new Drawing(
                     new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture"
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                             new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                         new DocumentFormat.OpenXml.Drawing.Graphic(
                             new DocumentFormat.OpenXml.Drawing.GraphicData(
                                 new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                     new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                         new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                         new DocumentFormat.OpenXml.Drawing.Blip() { Embed = relationshipId },
                                         new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                         new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                         new DocumentFormat.OpenXml.Drawing.Transform2D(
                                             new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                             new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000L }),
                                             new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                                 new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                             )
                                             { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                         )
                                     )
                                 )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                         )
                     )
                 );

            body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(element)));
        }
    }

    public class ImageExtractionListener : IEventListener
    {
        public List<(string Base64, int Width, int Height)> ExtractedImages { get; } = new();

        public void EventOccurred(IEventData data, EventType type)
        {
            if (data is ImageRenderInfo renderInfo)
            {
                var imgObject = renderInfo.GetImage();
                if (imgObject != null)
                {
                    var bytes = imgObject.GetImageBytes();
                    if (bytes?.Length > 0)
                    {
                        using var ms = new MemoryStream(bytes);
                        using var bmp = new Bitmap(ms);
                        using var outStream = new MemoryStream();
                        bmp.Save(outStream, ImageFormat.Png);

                        ExtractedImages.Add(
                            (Convert.ToBase64String(outStream.ToArray()), bmp.Width, bmp.Height)
                        );
                    }
                }
            }
        }

        public ICollection<EventType> GetSupportedEvents()
            => new[] { EventType.RENDER_IMAGE };


    }

    public class LocationTextExtractionStrategyEx : LocationTextExtractionStrategy
{
    private readonly List<TextChunkEx> _chunks = new();

    public override void EventOccurred(IEventData data, EventType type)
    {
        base.EventOccurred(data, type);

        if (type == EventType.RENDER_TEXT)
        {
            var renderInfo = (TextRenderInfo)data;
            var font = renderInfo.GetFont();
            var fontName = font != null ? font.GetFontProgram().ToString() : "Unknown";

            var rect = renderInfo.GetBaseline().GetBoundingRectangle();

            _chunks.Add(new TextChunkEx
            {
                Text = renderInfo.GetText(),
                FontSize = renderInfo.GetFontSize(),
                FontName = fontName,
                X = rect.GetX(),
                Y = rect.GetY(),
                Width = rect.GetWidth(),
                Height = rect.GetHeight(),
                IsBold = fontName.ToLower().Contains("bold"),
                IsItalic = fontName.ToLower().Contains("italic")
            });
        }
    }

    public List<TextChunkEx> GetChunks() => _chunks;
}

public class TextChunkEx
{
    public string Text { get; set; }
    public float FontSize { get; set; }
    public string FontName { get; set; }
    public float X { get; set; }
    public float Y { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
}

}

