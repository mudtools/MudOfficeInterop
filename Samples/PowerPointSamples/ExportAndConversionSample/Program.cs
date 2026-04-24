//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;

namespace ExportAndConversionSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - 导出与转换示例");

            Console.WriteLine("\n=== 示例1: 另存为PDF ===");
            SaveAsPdfDemo();

            Console.WriteLine("\n=== 示例2: 另存为不同格式 ===");
            SaveAsDifferentFormatsDemo();

            Console.WriteLine("\n=== 示例3: 导出幻灯片为图片 ===");
            ExportSlidesAsImagesDemo();

            Console.WriteLine("\n=== 示例4: 导出单个形状为图片 ===");
            ExportShapeAsImageDemo();

            Console.WriteLine("\n=== 示例5: 提取PPT纯文本内容 ===");
            ExtractTextContentDemo();

            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void SaveAsPdfDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "PDF导出示例";

                if (slide1?.Shapes?.Placeholders != null && slide1.Shapes.Placeholders.Count >= 2)
                {
                    var subtitle = slide1.Shapes.Placeholders[2];
                    if (subtitle?.TextFrame?.TextRange != null)
                        subtitle.TextFrame.TextRange.Text = "此演示文稿将被导出为PDF格式";
                }

                var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
                if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide2.Shapes.Title.TextFrame.TextRange.Text = "第二页内容";

                if (slide2?.Shapes?.Placeholders != null && slide2.Shapes.Placeholders.Count >= 2)
                {
                    var body = slide2.Shapes.Placeholders[2];
                    if (body?.TextFrame?.TextRange != null)
                        body.TextFrame.TextRange.Text = "这是PDF文档的第二页\n包含更多内容";
                }

                string tempDirectory = GetTempDirectory();
                string pdfPath = Path.Combine(tempDirectory, "ExportExample.pdf");

                presentation.SaveAs(pdfPath, PpSaveAsFileType.ppSaveAsPDF);
                Console.WriteLine($"演示文稿已另存为PDF: {pdfPath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"另存为PDF出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SaveAsDifferentFormatsDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide.Shapes.Title.TextFrame.TextRange.Text = "格式转换示例";

                string tempDirectory = GetTempDirectory();

                string pptxPath = Path.Combine(tempDirectory, "FormatExample.pptx");
                presentation.SaveAs(pptxPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"已保存为PPTX: {pptxPath}");

                string ppsxPath = Path.Combine(tempDirectory, "FormatExample.ppsx");
                presentation.SaveAs(ppsxPath, PpSaveAsFileType.ppSaveAsOpenXMLShow);
                Console.WriteLine($"已保存为PPSX(放映格式): {ppsxPath}");

                string potxPath = Path.Combine(tempDirectory, "FormatExample.potx");
                presentation.SaveAs(potxPath, PpSaveAsFileType.ppSaveAsOpenXMLTemplate);
                Console.WriteLine($"已保存为POTX(模板格式): {potxPath}");

                string xpsPath = Path.Combine(tempDirectory, "FormatExample.xps");
                presentation.SaveAs(xpsPath, PpSaveAsFileType.ppSaveAsXPS);
                Console.WriteLine($"已保存为XPS: {xpsPath}");

                string rtfPath = Path.Combine(tempDirectory, "FormatExample.rtf");
                presentation.SaveAs(rtfPath, PpSaveAsFileType.ppSaveAsRTF);
                Console.WriteLine($"已保存为RTF: {rtfPath}");

                Console.WriteLine("\n支持的常用格式:");
                Console.WriteLine("  ppSaveAsOpenXMLPresentation (24) - .pptx");
                Console.WriteLine("  ppSaveAsOpenXMLShow (28) - .ppsx");
                Console.WriteLine("  ppSaveAsOpenXMLTemplate (26) - .potx");
                Console.WriteLine("  ppSaveAsPDF (32) - .pdf");
                Console.WriteLine("  ppSaveAsXPS (33) - .xps");
                Console.WriteLine("  ppSaveAsRTF (6) - .rtf");
                Console.WriteLine("  ppSaveAsJPG (17) - .jpg");
                Console.WriteLine("  ppSaveAsPNG (18) - .png");
                Console.WriteLine("  ppSaveAsGIF (16) - .gif");
                Console.WriteLine("  ppSaveAsBMP (19) - .bmp");
                Console.WriteLine("  ppSaveAsTIF (21) - .tif");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"格式转换出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void ExportSlidesAsImagesDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "图片导出示例 - 第1页";

                var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide2.Shapes.Title.TextFrame.TextRange.Text = "图片导出示例 - 第2页";

                var slide3 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide3?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide3.Shapes.Title.TextFrame.TextRange.Text = "图片导出示例 - 第3页";

                string tempDirectory = GetTempDirectory();

                string pngFolder = Path.Combine(tempDirectory, "SlideImages_PNG");
                if (!Directory.Exists(pngFolder))
                    Directory.CreateDirectory(pngFolder);

                presentation.Export(pngFolder, "PNG", 1920, 1080);
                Console.WriteLine($"所有幻灯片已导出为PNG图片: {pngFolder}");

                string jpgFolder = Path.Combine(tempDirectory, "SlideImages_JPG");
                if (!Directory.Exists(jpgFolder))
                    Directory.CreateDirectory(jpgFolder);

                presentation.Export(jpgFolder, "JPG", 1280, 720);
                Console.WriteLine($"所有幻灯片已导出为JPG图片: {jpgFolder}");

                Console.WriteLine("\n提示: Export 方法将每页幻灯片导出为单独的图片文件");
                Console.WriteLine("  文件命名格式: Slide1.PNG, Slide2.PNG, ...");
                Console.WriteLine("  scaleWidth/scaleHeight 参数控制导出图片的分辨率");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出幻灯片为图片出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void ExportShapeAsImageDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var shape = shapes.AddShape(MsoAutoShapeType.msoShapeHeart, 200, 100, 300, 250);
                if (shape != null)
                {
                    var fill = shape.Fill;
                    fill?.Solid();
                    if (fill?.ForeColor != null)
                        fill.ForeColor.RGB = System.Drawing.Color.Red;

                    string tempDirectory = GetTempDirectory();
                    string shapePngPath = Path.Combine(tempDirectory, "HeartShape.png");
                    shape.Export(shapePngPath, PpShapeFormat.ppShapeFormatPNG, 300, 250);
                    Console.WriteLine($"心形形状已导出为PNG: {shapePngPath}");

                    string shapeJpgPath = Path.Combine(tempDirectory, "HeartShape.jpg");
                    shape.Export(shapeJpgPath, PpShapeFormat.ppShapeFormatJPG, 300, 250);
                    Console.WriteLine($"心形形状已导出为JPG: {shapeJpgPath}");

                    Console.WriteLine("\n支持的形状导出格式:");
                    Console.WriteLine("  ppShapeFormatGIF - GIF格式");
                    Console.WriteLine("  ppShapeFormatJPG - JPG格式");
                    Console.WriteLine("  ppShapeFormatPNG - PNG格式");
                    Console.WriteLine("  ppShapeFormatBMP - BMP格式");
                    Console.WriteLine("  ppShapeFormatWMF - WMF格式");
                    Console.WriteLine("  ppShapeFormatEMF - EMF格式");
                }

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出形状为图片出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void ExtractTextContentDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "内容提取示例";

                if (slide1?.Shapes?.Placeholders != null && slide1.Shapes.Placeholders.Count >= 2)
                {
                    var subtitle = slide1.Shapes.Placeholders[2];
                    if (subtitle?.TextFrame?.TextRange != null)
                        subtitle.TextFrame.TextRange.Text = "副标题文本";
                }

                var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
                if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide2.Shapes.Title.TextFrame.TextRange.Text = "详细内容";

                if (slide2?.Shapes?.Placeholders != null && slide2.Shapes.Placeholders.Count >= 2)
                {
                    var body = slide2.Shapes.Placeholders[2];
                    if (body?.TextFrame?.TextRange != null)
                        body.TextFrame.TextRange.Text = "第一点：文本提取功能\n第二点：搜索引擎索引\n第三点：内容审核";
                }

                var slide3 = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);
                var shapes3 = slide3?.Shapes;
                if (shapes3 != null)
                {
                    var textbox = shapes3.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 400, 40);
                    if (textbox?.TextFrame?.TextRange != null)
                        textbox.TextFrame.TextRange.Text = "空白幻灯片上的独立文本框";
                }

                Console.WriteLine("--- 提取的文本内容 ---");
                var allText = ExtractAllText(presentation);
                Console.WriteLine(allText);
                Console.WriteLine("--- 提取结束 ---");

                string tempDirectory = GetTempDirectory();
                string textFilePath = Path.Combine(tempDirectory, "ExtractedContent.txt");
                File.WriteAllText(textFilePath, allText);
                Console.WriteLine($"\n提取的文本已保存到: {textFilePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"内容提取出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void CompleteExampleWithHelpers()
        {
            try
            {
                var converter = new PowerPointConverter();

                string pptxPath = converter.CreateSamplePresentation();
                Console.WriteLine($"已创建示例演示文稿: {pptxPath}");

                string pdfPath = converter.ConvertToPdf(pptxPath);
                Console.WriteLine($"已转换为PDF: {pdfPath}");

                string imageFolder = converter.ConvertToImages(pptxPath, "PNG");
                Console.WriteLine($"已转换为图片: {imageFolder}");

                string textContent = converter.ExtractText(pptxPath);
                Console.WriteLine($"提取的文本内容长度: {textContent.Length} 字符");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"辅助类完整示例出错: {ex.Message}");
            }
        }

        static string ExtractAllText(IPowerPointPresentation presentation)
        {
            var sb = new System.Text.StringBuilder();

            var slides = presentation.GetAllSlides();
            foreach (var slide in slides)
            {
                sb.AppendLine($"=== 幻灯片 {slide.SlideIndex} ===");

                var shapes = slide.Shapes;
                if (shapes != null)
                {
                    foreach (var shape in shapes)
                    {
                        if (shape != null && shape.HasTextFrame)
                        {
                            var textRange = shape.TextFrame?.TextRange;
                            if (textRange != null && !string.IsNullOrEmpty(textRange.Text))
                            {
                                sb.AppendLine($"  [{shape.Name}]: {textRange.Text}");
                            }
                        }

                        if (shape != null && shape.HasTable)
                        {
                            var table = shape.Table;
                            if (table != null)
                            {
                                sb.AppendLine($"  [表格 {table.Rows?.Count ?? 0}×{table.Columns?.Count ?? 0}]:");
                                var rows = table.Rows;
                                var cols = table.Columns;
                                if (rows != null && cols != null)
                                {
                                    for (int r = 1; r <= rows.Count; r++)
                                    {
                                        for (int c = 1; c <= cols.Count; c++)
                                        {
                                            var cell = table.Cell(r, c);
                                            if (cell?.Shape?.TextFrame?.TextRange != null)
                                            {
                                                sb.Append($"  {cell.Shape.TextFrame.TextRange.Text}");
                                                if (c < cols.Count) sb.Append(" | ");
                                            }
                                        }
                                        sb.AppendLine();
                                    }
                                }
                            }
                        }
                    }
                }

                sb.AppendLine();
                slide.Dispose();
            }

            return sb.ToString();
        }

        static string GetTempDirectory()
        {

            string tempDirectory = Path.Combine(AppContext.BaseDirectory, "Output\\PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }

    public class PowerPointConverter
    {
        public string CreateSamplePresentation()
        {
            using var app = PowerPointFactory.BlankDocument();
            var presentation = app.ActivePresentation;

            var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
            if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                slide1.Shapes.Title.TextFrame.TextRange.Text = "转换示例";

            if (slide1?.Shapes?.Placeholders != null && slide1.Shapes.Placeholders.Count >= 2)
            {
                var subtitle = slide1.Shapes.Placeholders[2];
                if (subtitle?.TextFrame?.TextRange != null)
                    subtitle.TextFrame.TextRange.Text = "此演示文稿将用于格式转换";
            }

            var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
            if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                slide2.Shapes.Title.TextFrame.TextRange.Text = "内容页";

            if (slide2?.Shapes?.Placeholders != null && slide2.Shapes.Placeholders.Count >= 2)
            {
                var body = slide2.Shapes.Placeholders[2];
                if (body?.TextFrame?.TextRange != null)
                    body.TextFrame.TextRange.Text = "这是内容页的文本\n用于测试文本提取功能";
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);

            string filePath = Path.Combine(tempDirectory, $"ConverterSample_{Guid.NewGuid():N}.pptx");
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();

            return filePath;
        }

        public string ConvertToPdf(string pptxPath)
        {
            using var app = PowerPointFactory.Open(pptxPath);
            var presentation = app.ActivePresentation;

            string pdfPath = Path.ChangeExtension(pptxPath, ".pdf");
            presentation.SaveAs(pdfPath, PpSaveAsFileType.ppSaveAsPDF);
            presentation.Close();

            return pdfPath;
        }

        public string ConvertToImages(string pptxPath, string format = "PNG")
        {
            using var app = PowerPointFactory.Open(pptxPath);
            var presentation = app.ActivePresentation;

            string imageFolder = Path.Combine(
                Path.GetDirectoryName(pptxPath)!,
                Path.GetFileNameWithoutExtension(pptxPath) + "_Images"
            );

            if (!Directory.Exists(imageFolder))
                Directory.CreateDirectory(imageFolder);

            presentation.Export(imageFolder, format, 1920, 1080);
            presentation.Close();

            return imageFolder;
        }

        public string ExtractText(string pptxPath)
        {
            using var app = PowerPointFactory.Open(pptxPath);
            var presentation = app.ActivePresentation;

            var sb = new System.Text.StringBuilder();
            var slides = presentation.GetAllSlides();
            foreach (var slide in slides)
            {
                var shapes = slide.Shapes;
                if (shapes != null)
                {
                    foreach (var shape in shapes)
                    {
                        if (shape != null && shape.HasTextFrame)
                        {
                            var text = shape.TextFrame?.TextRange?.Text;
                            if (!string.IsNullOrEmpty(text))
                                sb.AppendLine(text);
                        }
                    }
                }
                slide.Dispose();
            }

            presentation.Close();
            return sb.ToString();
        }
    }
}
