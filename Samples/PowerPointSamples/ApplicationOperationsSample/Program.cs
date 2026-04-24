using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;

namespace ApplicationOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - 应用程序级操作示例");

            Console.WriteLine("\n=== 示例1: 启动与关闭应用程序 ===");
            StartAndQuitDemo();

            Console.WriteLine("\n=== 示例2: 应用程序可见性控制 ===");
            VisibilityControlDemo();

            Console.WriteLine("\n=== 示例3: 演示文稿管理 - 新建演示文稿 ===");
            CreatePresentationDemo();

            Console.WriteLine("\n=== 示例4: 演示文稿管理 - 打开现有文件 ===");
            OpenPresentationDemo();

            Console.WriteLine("\n=== 示例5: 应用程序基础属性 ===");
            ApplicationPropertiesDemo();

            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void StartAndQuitDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();

                app.Visible = true;
                Console.WriteLine("PowerPoint 应用程序已启动并设为可见");

                Console.WriteLine("应用程序将在3秒后关闭...");
                Thread.Sleep(3000);

                app.Quit();
                Console.WriteLine("应用程序已通过 Quit() 关闭");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"启动与关闭应用程序出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void VisibilityControlDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();

                app.Visible = false;
                Console.WriteLine("应用程序设为不可见（后台运行模式） - 适用于自动化任务");

                Thread.Sleep(1000);

                app.Visible = true;
                Console.WriteLine("应用程序已设为可见 - 适用于调试和交互式操作");

                Console.WriteLine($"当前可见性: {app.Visible}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"可见性控制出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void CreatePresentationDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                if (presentation != null)
                {
                    Console.WriteLine($"新建演示文稿名称: {presentation.Name}");
                    Console.WriteLine($"幻灯片数量: {presentation.SlideCount}");
                }

                string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                string filePath = Path.Combine(tempDirectory, "NewPresentation.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"新建演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"新建演示文稿出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void OpenPresentationDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
                string filePath = Path.Combine(tempDirectory, "NewPresentation.pptx");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("请先运行示例3创建演示文稿文件");
                    return;
                }

                app = PowerPointFactory.Open(filePath);
                app.Visible = true;

                var presentation = app.ActivePresentation;
                if (presentation != null)
                {
                    Console.WriteLine($"已打开演示文稿: {presentation.Name}");
                    Console.WriteLine($"文件路径: {presentation.FullName}");
                    Console.WriteLine($"幻灯片数量: {presentation.SlideCount}");
                    Console.WriteLine($"是否只读: {presentation.ReadOnly}");
                }

                Thread.Sleep(2000);
                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打开演示文稿出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void ApplicationPropertiesDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                Console.WriteLine($"应用程序名称: {app.Name}");
                Console.WriteLine($"应用程序版本: {app.Version}");
                Console.WriteLine($"安装路径: {app.Path}");
                Console.WriteLine($"是否激活: {app.IsActive}");

                app.WindowStateValue = (int)PpWindowState.ppWindowMaximized;
                Console.WriteLine("窗口已最大化");

                Console.WriteLine($"窗口位置 - 左: {app.Left}, 上: {app.Top}");
                Console.WriteLine($"窗口尺寸 - 宽: {app.Width}, 高: {app.Height}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用程序属性操作出错: {ex.Message}");
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
                var manager = new PowerPointAppManager();

                string filePath = manager.CreateAndSavePresentation("应用程序操作示例", "这是一个通过MudTools创建的演示文稿");
                Console.WriteLine($"辅助类创建的演示文稿: {filePath}");

                manager.OpenAndModifyPresentation(filePath, "\n\n修改时间: " + DateTime.Now.ToString());
                Console.WriteLine("演示文稿已修改并保存");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"辅助类完整示例出错: {ex.Message}");
            }
        }
    }

    public class PowerPointAppManager
    {
        public string CreateAndSavePresentation(string title, string content)
        {
            using var app = PowerPointFactory.BlankDocument();
            var presentation = app.ActivePresentation;

            var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
            if (slide != null)
            {
                var shapes = slide.Shapes;
                if (shapes != null)
                {
                    var titleShape = shapes.Title;
                    if (titleShape?.TextFrame?.TextRange != null)
                    {
                        titleShape.TextFrame.TextRange.Text = title;
                    }

                    var bodyPlaceholder = shapes.Placeholders?[2];
                    if (bodyPlaceholder?.TextFrame?.TextRange != null)
                    {
                        bodyPlaceholder.TextFrame.TextRange.Text = content;
                    }
                }
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
            {
                Directory.CreateDirectory(tempDirectory);
            }

            string filePath = Path.Combine(tempDirectory, $"AppManager_{Guid.NewGuid():N}.pptx");
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();

            return filePath;
        }

        public void OpenAndModifyPresentation(string filePath, string additionalContent)
        {
            using var app = PowerPointFactory.Open(filePath);
            var presentation = app.ActivePresentation;

            var slide = presentation.GetSlide(1);
            if (slide != null)
            {
                var shapes = slide.Shapes;
                if (shapes != null)
                {
                    var bodyPlaceholder = shapes.Placeholders?[2];
                    if (bodyPlaceholder?.TextFrame?.TextRange != null)
                    {
                        bodyPlaceholder.TextFrame.TextRange.InsertAfter(additionalContent);
                    }
                }
            }

            presentation.Save();
            presentation.Close();
        }
    }
}
