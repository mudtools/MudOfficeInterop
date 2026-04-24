//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;

namespace SlideShowControlSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - 幻灯片放映控制示例");

            Console.WriteLine("\n=== 示例1: 基本放映设置 ===");
            BasicSlideShowSettingsDemo();

            Console.WriteLine("\n=== 示例2: 展台自动循环放映 ===");
            KioskSlideShowDemo();

            Console.WriteLine("\n=== 示例3: 自定义放映范围 ===");
            CustomSlideShowRangeDemo();

            Console.WriteLine("\n=== 示例4: 监听幻灯片放映事件 ===");
            SlideShowEventsDemo();

            Console.WriteLine("\n=== 示例5: 放映窗口控制 ===");
            SlideShowWindowDemo();

            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void BasicSlideShowSettingsDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                for (int i = 1; i <= 5; i++)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = $"幻灯片 {i}";

                    if (slide?.Shapes?.Placeholders != null && slide.Shapes.Placeholders.Count >= 2)
                    {
                        var subtitle = slide.Shapes.Placeholders[2];
                        if (subtitle?.TextFrame?.TextRange != null)
                            subtitle.TextFrame.TextRange.Text = $"这是第 {i} 页内容";
                    }
                }

                var settings = presentation.SlideShowSettings;
                if (settings != null)
                {
                    Console.WriteLine($"当前放映类型: {settings.ShowType}");
                    Console.WriteLine($"起始幻灯片: {settings.StartingSlide}");
                    Console.WriteLine($"结束幻灯片: {settings.EndingSlide}");
                    Console.WriteLine($"是否循环放映: {settings.LoopUntilStopped}");

                    settings.ShowType = PpSlideShowType.ppShowTypeSpeaker;
                    Console.WriteLine("\n已设置放映类型为: 演讲者放映（全屏幕）");

                    Console.WriteLine("\n支持的放映类型:");
                    Console.WriteLine("  ppShowTypeSpeaker (1) - 演讲者放映（全屏幕）");
                    Console.WriteLine("  ppShowTypeWindow (2) - 观众自行浏览（窗口）");
                    Console.WriteLine("  ppShowTypeKiosk (3) - 展台浏览（全屏幕）");
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "BasicSlideShow.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"基本放映设置出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void KioskSlideShowDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                for (int i = 1; i <= 4; i++)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = $"展台展示 - 第{i}页";

                    if (slide?.Shapes?.Placeholders != null && slide.Shapes.Placeholders.Count >= 2)
                    {
                        var subtitle = slide.Shapes.Placeholders[2];
                        if (subtitle?.TextFrame?.TextRange != null)
                            subtitle.TextFrame.TextRange.Text = $"自动循环放映内容 {i}";
                    }
                }

                var settings = presentation.SlideShowSettings;
                if (settings != null)
                {
                    settings.ShowType = PpSlideShowType.ppShowTypeKiosk;
                    settings.LoopUntilStopped = true;
                    settings.AdvanceMode = PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;

                    Console.WriteLine("已设置展台自动循环放映:");
                    Console.WriteLine("  放映类型: 展台浏览（全屏幕）");
                    Console.WriteLine("  循环放映: 是");
                    Console.WriteLine("  切换方式: 使用幻灯片计时");
                    Console.WriteLine("\n提示: 展台模式下，ESC键退出放映，幻灯片会自动循环");
                }

                var slide1 = presentation.GetSlide(1);
                var slideShowTransition = slide1?.SlideShowTransition;
                if (slideShowTransition != null)
                {
                    slideShowTransition.AdvanceOnTime = true;
                    slideShowTransition.AdvanceTime = 3;
                    Console.WriteLine($"第1张幻灯片: 自动切换间隔设为 3 秒");
                }

                var slide2 = presentation.GetSlide(2);
                var transition2 = slide2?.SlideShowTransition;
                if (transition2 != null)
                {
                    transition2.AdvanceOnTime = true;
                    transition2.AdvanceTime = 5;
                    Console.WriteLine($"第2张幻灯片: 自动切换间隔设为 5 秒");
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "KioskSlideShow.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"展台自动循环放映出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void CustomSlideShowRangeDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                string[] slideTitles = { "封面", "目录", "内容A", "内容B", "内容C", "总结" };
                foreach (var title in slideTitles)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = title;
                }

                var settings = presentation.SlideShowSettings;
                if (settings != null)
                {
                    settings.StartingSlide = 2;
                    settings.EndingSlide = 5;
                    Console.WriteLine("已设置放映范围: 从第2张到第5张");
                    Console.WriteLine("  跳过封面和总结页，只放映目录和内容页");

                    settings.ShowType = PpSlideShowType.ppShowTypeSpeaker;
                    settings.LoopUntilStopped = false;
                }

                Console.WriteLine("\n放映范围设置说明:");
                Console.WriteLine("  StartingSlide - 起始幻灯片编号");
                Console.WriteLine("  EndingSlide - 结束幻灯片编号");
                Console.WriteLine("  NamedShow - 使用已命名的自定义放映");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "CustomRangeSlideShow.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自定义放映范围出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SlideShowEventsDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                for (int i = 1; i <= 3; i++)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = $"事件测试 - 第{i}页";
                }

                app.SlideShowNextSlide += (window) =>
                {
                    try
                    {
                        if (window != null)
                        {
                            var currentSlide = window.View?.Slide;
                            if (currentSlide != null)
                            {
                                Console.WriteLine($"  [事件] 切换到幻灯片: 索引={currentSlide.SlideIndex}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  [事件处理出错] {ex.Message}");
                    }
                };

                Console.WriteLine("已订阅 SlideShowNextSlide 事件");
                Console.WriteLine("当幻灯片放映翻页时，将触发 .NET 代码回调");
                Console.WriteLine("\n提示: COM 事件订阅是 PowerPoint 自动化的高级功能");
                Console.WriteLine("  常用事件包括:");
                Console.WriteLine("  - SlideShowNextSlide: 翻到下一张幻灯片时触发");
                Console.WriteLine("  - SlideShowBegin: 放映开始时触发");
                Console.WriteLine("  - SlideShowEnd: 放映结束时触发");
                Console.WriteLine("\n启动放映后，每次翻页都会触发事件回调...");
                Console.WriteLine("（此示例仅注册事件，不自动启动放映）");

                var settings = presentation.SlideShowSettings;
                if (settings != null)
                {
                    settings.ShowType = PpSlideShowType.ppShowTypeSpeaker;
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideShowEvents.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"监听放映事件出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SlideShowWindowDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                for (int i = 1; i <= 3; i++)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                    if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                        slide.Shapes.Title.TextFrame.TextRange.Text = $"放映窗口控制 - 第{i}页";
                }

                var settings = presentation.SlideShowSettings;
                if (settings != null)
                {
                    settings.ShowType = PpSlideShowType.ppShowTypeWindow;
                    Console.WriteLine("已设置放映类型为: 观众自行浏览（窗口模式）");
                    Console.WriteLine("  窗口模式下放映不会全屏，用户可以自由浏览");
                }

                Console.WriteLine("\n放映窗口操作说明:");
                Console.WriteLine("  SlideShowWindow.View - 放映视图控制");
                Console.WriteLine("  View.GotoSlide(n) - 跳转到指定幻灯片");
                Console.WriteLine("  View.Next() - 下一张");
                Console.WriteLine("  View.Previous() - 上一张");
                Console.WriteLine("  View.Exit() - 退出放映");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideShowWindow.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"放映窗口控制出错: {ex.Message}");
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
                var controller = new SlideShowController();

                string filePath = controller.CreatePresentationWithTransitions();
                Console.WriteLine($"已创建带切换效果的演示文稿: {filePath}");

                controller.ConfigureKioskShow(filePath);
                Console.WriteLine("已配置展台循环放映模式");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"辅助类完整示例出错: {ex.Message}");
            }
        }

        static string GetTempDirectory()
        {
            string tempDirectory = Path.Combine(AppContext.BaseDirectory, "Output\\PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }

    public class SlideShowController
    {
        public string CreatePresentationWithTransitions()
        {
            using var app = PowerPointFactory.BlankDocument();
            var presentation = app.ActivePresentation;

            var slide1 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
            if (slide1?.Shapes?.Title?.TextFrame?.TextRange != null)
                slide1.Shapes.Title.TextFrame.TextRange.Text = "自动展示";

            if (slide1?.Shapes?.Placeholders != null && slide1.Shapes.Placeholders.Count >= 2)
            {
                var subtitle = slide1.Shapes.Placeholders[2];
                if (subtitle?.TextFrame?.TextRange != null)
                    subtitle.TextFrame.TextRange.Text = "展台模式自动循环放映";
            }

            var transition1 = slide1?.SlideShowTransition;
            if (transition1 != null)
            {
                transition1.AdvanceOnTime = true;
                transition1.AdvanceTime = 3;
                transition1.EntryEffect = PpEntryEffect.ppEffectFade;
            }

            var slide2 = presentation.AddSlide(PpSlideLayout.ppLayoutText);
            if (slide2?.Shapes?.Title?.TextFrame?.TextRange != null)
                slide2.Shapes.Title.TextFrame.TextRange.Text = "产品特性";

            if (slide2?.Shapes?.Placeholders != null && slide2.Shapes.Placeholders.Count >= 2)
            {
                var body = slide2.Shapes.Placeholders[2];
                if (body?.TextFrame?.TextRange != null)
                    body.TextFrame.TextRange.Text = "特性一：高性能\n特性二：易扩展\n特性三：安全可靠";
            }

            var transition2 = slide2?.SlideShowTransition;
            if (transition2 != null)
            {
                transition2.AdvanceOnTime = true;
                transition2.AdvanceTime = 5;
                transition2.EntryEffect = PpEntryEffect.ppEffectPushDown;
            }

            var slide3 = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
            if (slide3?.Shapes?.Title?.TextFrame?.TextRange != null)
                slide3.Shapes.Title.TextFrame.TextRange.Text = "谢谢观看";

            var transition3 = slide3?.SlideShowTransition;
            if (transition3 != null)
            {
                transition3.AdvanceOnTime = true;
                transition3.AdvanceTime = 3;
                transition3.EntryEffect = PpEntryEffect.ppEffectFade;
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);

            string filePath = Path.Combine(tempDirectory, $"SlideShowController_{Guid.NewGuid():N}.pptx");
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();

            return filePath;
        }

        public void ConfigureKioskShow(string filePath)
        {
            using var app = PowerPointFactory.Open(filePath);
            var presentation = app.ActivePresentation;

            var settings = presentation.SlideShowSettings;
            if (settings != null)
            {
                settings.ShowType = PpSlideShowType.ppShowTypeKiosk;
                settings.LoopUntilStopped = true;
                settings.AdvanceMode = PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;
            }

            presentation.Save();
            presentation.Close();
        }
    }
}
