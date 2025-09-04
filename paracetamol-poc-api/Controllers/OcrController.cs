using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfiumViewer;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Tesseract;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using PdfiumDocument = PdfiumViewer.PdfDocument;
using PdfpigDocument = UglyToad.PdfPig.PdfDocument;

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
                    using (var pageImg = pdfDoc.Render(i, 400, 400, true)) // high DPI for small text
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
}

