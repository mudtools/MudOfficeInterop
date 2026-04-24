//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;

namespace PresentationAndSlideSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - Presentation 与 Slide 操作示例");

            Console.WriteLine("\n=== 示例1: 遍历所有幻灯片 ===");
            IterateSlidesDemo();

            Console.WriteLine("\n=== 示例2: 插入不同版式的新幻灯片 ===");
            AddSlidesWithLayoutsDemo();

            Console.WriteLine("\n=== 示例3: 幻灯片删除与复制 ===");
            SlideDeleteAndDuplicateDemo();

            Console.WriteLine("\n=== 示例4: 幻灯片移动与排序 ===");
            SlideMoveAndOrderDemo();

            Console.WriteLine("\n=== 示例5: 页面设置操作 ===");
            PageSetupDemo();

            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void IterateSlidesDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "第一页幻灯片";

                var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
                if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide2.Shapes.Title.TextFrame.TextRange.Text = "第二页幻灯片";

                var slide3 = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);
                var shapes3 = slide3?.Shapes;
                if (shapes3 != null)
                {
                    var textbox = shapes3.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 500, 50);
                    if (textbox?.TextFrame?.TextRange != null)
                        textbox.TextFrame.TextRange.Text = "第三页 - 空白版式幻灯片";
                }

                Console.WriteLine($"演示文稿共有 {presentation.SlideCount} 张幻灯片");
                Console.WriteLine("--- 通过索引遍历 ---");
                for (int i = 1; i <= presentation.SlideCount; i++)
                {
                    using var slide = presentation.GetSlide(i);
                    if (slide != null)
                    {
                        Console.WriteLine($"  幻灯片 {i}: 索引={slide.SlideIndex}, 编号={slide.SlideNumber}, 版式={slide.Layout}, ID={slide.SlideID}");
                    }
                }

                Console.WriteLine("--- 通过 GetAllSlides 遍历 ---");
                var allSlides = presentation.GetAllSlides();
                foreach (var slide in allSlides)
                {
                    Console.WriteLine($"  幻灯片: 索引={slide.SlideIndex}, 版式={slide.Layout}");
                    slide.Dispose();
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "IterateSlides.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历幻灯片出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void AddSlidesWithLayoutsDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var layouts = new (PpSlideLayout Layout, string Title)[]
                {
                    (PpSlideLayout.ppLayoutTitle, "标题幻灯片"),
                    (PpSlideLayout.ppLayoutText, "文本幻灯片"),
                    (PpSlideLayout.ppLayoutBlank, "空白幻灯片"),
                    (PpSlideLayout.ppLayoutTwoColumnText, "两栏文本幻灯片"),
                    (PpSlideLayout.ppLayoutTable, "表格幻灯片"),
                    (PpSlideLayout.ppLayoutChart, "图表幻灯片"),
                    (PpSlideLayout.ppLayoutTitleOnly, "仅标题幻灯片"),
                    (PpSlideLayout.ppLayoutSectionHeader, "节标题幻灯片"),
                };

                foreach (var (layout, title) in layouts)
                {
                    var slide = presentation.AddSlide(layout);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                    {
                        slide.Shapes.Title.TextFrame.TextRange.Text = title;
                    }
                    Console.WriteLine($"已添加: {title} (版式={layout})");
                }

                Console.WriteLine($"\n共添加 {presentation.SlideCount} 张不同版式的幻灯片");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideLayouts.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入不同版式幻灯片出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SlideDeleteAndDuplicateDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "原始幻灯片";

                var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
                if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide2.Shapes.Title.TextFrame.TextRange.Text = "第二张幻灯片";

                Console.WriteLine($"初始幻灯片数量: {presentation.SlideCount}");

                var duplicatedRange = slide1.Duplicate();
                Console.WriteLine($"复制幻灯片后数量: {presentation.SlideCount}");

                presentation.RemoveSlide(3);
                Console.WriteLine($"删除第3张幻灯片后数量: {presentation.SlideCount}");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideDeleteDuplicate.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"幻灯片删除与复制出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SlideMoveAndOrderDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                string[] slideTitles = { "幻灯片 A", "幻灯片 B", "幻灯片 C", "幻灯片 D" };
                foreach (var title in slideTitles)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = title;
                }

                Console.WriteLine("移动前顺序:");
                PrintSlideOrder(presentation);

                var slideToMove = presentation.GetSlide(4);
                slideToMove?.MoveTo(1);
                Console.WriteLine("\n将第4张幻灯片移动到第1位后:");
                PrintSlideOrder(presentation);

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideMoveOrder.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"幻灯片移动与排序出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void PageSetupDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "页面设置示例";

                var master = slide1?.Master;
                if (master != null)
                {
                    Console.WriteLine($"幻灯片母版宽度: {master.Width} 磅");
                    Console.WriteLine($"幻灯片母版高度: {master.Height} 磅");
                    Console.WriteLine($"标准16:9宽度约为960磅，高度约为540磅");
                    Console.WriteLine($"标准4:3宽度约为720磅，高度约为540磅");
                }

                Console.WriteLine("\n提示: PPT内部单位为磅(Point)，1英寸=72磅");
                Console.WriteLine("  厘米转磅: cm * 28.35");
                Console.WriteLine("  英寸转磅: inch * 72");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "PageSetup.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页面设置操作出错: {ex.Message}");
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
                var builder = new PresentationBuilder();

                string filePath = builder.BuildMultiSlidePresentation(
                    "季度汇报",
                    new[]
                    {
                        "第一季度业绩概述",
                        "第二季度业绩概述",
                        "第三季度业绩概述",
                        "第四季度业绩概述",
                        "年度总结"
                    }
                );
                Console.WriteLine($"辅助类创建的多幻灯片演示文稿: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"辅助类完整示例出错: {ex.Message}");
            }
        }

        static void PrintSlideOrder(IPowerPointPresentation presentation)
        {
            for (int i = 1; i <= presentation.SlideCount; i++)
            {
                using var slide = presentation.GetSlide(i);
                if (slide != null)
                {
                    var titleText = slide.Shapes?.Title?.TextFrame?.TextRange?.Text ?? "(无标题)";
                    Console.WriteLine($"  第{i}张: {titleText}");
                }
            }
        }

        static string GetTempDirectory()
        {
            string tempDirectory = Path.Combine(AppContext.BaseDirectory, "Output\\PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
            {
                Directory.CreateDirectory(tempDirectory);
            }
            return tempDirectory;
        }
    }

    public class PresentationBuilder
    {
        public string BuildMultiSlidePresentation(string mainTitle, string[] slideTitles)
        {
            using var app = PowerPointFactory.BlankDocument();
            var presentation = app.ActivePresentation;

            var titleSlide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
            if (titleSlide?.Shapes?.Title?.TextFrame?.TextRange != null)
                titleSlide.Shapes.Title.TextFrame.TextRange.Text = mainTitle;

            if (titleSlide?.Shapes?.Placeholders != null && titleSlide.Shapes.Placeholders.Count >= 2)
            {
                var subtitle = titleSlide.Shapes.Placeholders[2];
                if (subtitle?.TextFrame?.TextRange != null)
                    subtitle.TextFrame.TextRange.Text = $"共 {slideTitles.Length} 个章节 | 创建于 {DateTime.Now:yyyy-MM-dd}";
            }

            foreach (var title in slideTitles)
            {
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutSectionHeader);
                if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);

            string filePath = Path.Combine(tempDirectory, $"Presentation_{Guid.NewGuid():N}.pptx");
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();

            return filePath;
        }
    }
}
