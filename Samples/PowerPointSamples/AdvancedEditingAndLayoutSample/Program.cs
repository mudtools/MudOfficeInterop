//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;
using System.Drawing;

namespace AdvancedEditingAndLayoutSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - 高级编辑与排版示例");

            Console.WriteLine("\n=== 示例1: 数据合并 - 批量生成铭牌 ===");
            DataMergeNameplateDemo();

            Console.WriteLine("\n=== 示例2: 数据合并 - 批量生成奖状 ===");
            DataMergeCertificateDemo();

            Console.WriteLine("\n=== 示例3: 母版操作 - 访问 SlideMaster ===");
            SlideMasterDemo();

            Console.WriteLine("\n=== 示例4: 占位符操作 ===");
            PlaceholderDemo();

            Console.WriteLine("\n=== 示例5: 文本替换批量处理 ===");
            BatchReplaceTextDemo();

            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void DataMergeNameplateDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var nameplateData = new[]
                {
                    new { Name = "张三", Title = "首席执行官", Company = "星辰科技" },
                    new { Name = "李四", Title = "首席技术官", Company = "星辰科技" },
                    new { Name = "王五", Title = "产品总监", Company = "星辰科技" },
                    new { Name = "赵六", Title = "设计总监", Company = "星辰科技" },
                };

                foreach (var data in nameplateData)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);
                    if (slide == null) continue;

                    var shapes = slide.Shapes;

                    var bg = shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 100, 100, 500, 300);
                    if (bg != null)
                    {
                        var fill = bg.Fill;
                        fill?.Solid();
                        if (fill?.ForeColor != null)
                            fill.ForeColor.RGB = Color.FromArgb(240, 248, 255);

                        var line = bg.Line;
                        if (line != null)
                        {
                            line.Visible = true;
                            if (line.ForeColor != null)
                                line.ForeColor.RGB = Color.SteelBlue;
                            line.Weight = 2f;
                        }
                    }

                    var nameBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 150, 140, 400, 60);
                    if (nameBox?.TextFrame?.TextRange != null)
                    {
                        nameBox.TextFrame.TextRange.Text = data.Name;
                        nameBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = nameBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 36;
                            font.Bold = true;
                            font.Color.RGB = Color.DarkSlateBlue;
                        }
                    }

                    var titleBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 150, 220, 400, 40);
                    if (titleBox?.TextFrame?.TextRange != null)
                    {
                        titleBox.TextFrame.TextRange.Text = data.Title;
                        titleBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = titleBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 20;
                            font.Color.RGB = Color.Gray;
                        }
                    }

                    var companyBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 150, 290, 400, 40);
                    if (companyBox?.TextFrame?.TextRange != null)
                    {
                        companyBox.TextFrame.TextRange.Text = data.Company;
                        companyBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = companyBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 16;
                            font.Color.RGB = Color.SteelBlue;
                        }
                    }

                    Console.WriteLine($"  已生成铭牌: {data.Name} - {data.Title}");
                }

                Console.WriteLine($"共生成 {presentation.SlideCount} 张铭牌幻灯片");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "Nameplates.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"铭牌演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据合并铭牌出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void DataMergeCertificateDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var certificateData = new[]
                {
                    new { Name = "张三", Award = "年度最佳员工", Date = "2026年3月" },
                    new { Name = "李四", Award = "技术创新奖", Date = "2026年3月" },
                    new { Name = "王五", Award = "优秀团队领导", Date = "2026年3月" },
                };

                foreach (var data in certificateData)
                {
                    var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);
                    if (slide == null) continue;

                    var shapes = slide.Shapes;

                    var border = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 30, 30, 660, 450);
                    if (border != null)
                    {
                        var fill = border.Fill;
                        fill?.Solid();
                        if (fill?.ForeColor != null)
                            fill.ForeColor.RGB = Color.FromArgb(255, 253, 240);

                        var line = border.Line;
                        if (line != null)
                        {
                            line.Visible = true;
                            if (line.ForeColor != null)
                                line.ForeColor.RGB = Color.Gold;
                            line.Weight = 3f;
                        }
                    }

                    var innerBorder = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 45, 45, 630, 420);
                    if (innerBorder != null)
                    {
                        var fill = innerBorder.Fill;
                        fill?.Visible = false;

                        var line = innerBorder.Line;
                        if (line != null)
                        {
                            line.Visible = true;
                            if (line.ForeColor != null)
                                line.ForeColor.RGB = Color.Gold;
                            line.Weight = 1f;
                            line.DashStyle = MsoLineDashStyle.msoLineDash;
                        }
                    }

                    var certTitle = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 80, 560, 60);
                    if (certTitle?.TextFrame?.TextRange != null)
                    {
                        certTitle.TextFrame.TextRange.Text = "荣 誉 证 书";
                        certTitle.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = certTitle.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "华文行楷";
                            font.Size = 40;
                            font.Bold = true;
                            font.Color.RGB = Color.DarkRed;
                        }
                    }

                    var nameBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 180, 560, 50);
                    if (nameBox?.TextFrame?.TextRange != null)
                    {
                        nameBox.TextFrame.TextRange.Text = $"{data.Name} 同志";
                        nameBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = nameBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 28;
                            font.Bold = true;
                            font.Color.RGB = Color.Black;
                        }
                    }

                    var awardBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 250, 560, 80);
                    if (awardBox?.TextFrame?.TextRange != null)
                    {
                        awardBox.TextFrame.TextRange.Text = $"在2025年度工作中表现突出，荣获\n{data.Award}";
                        awardBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                        var font = awardBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 18;
                            font.Color.RGB = Color.DarkSlateGray;
                        }
                    }

                    var dateBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 380, 560, 40);
                    if (dateBox?.TextFrame?.TextRange != null)
                    {
                        dateBox.TextFrame.TextRange.Text = $"颁发日期: {data.Date}";
                        dateBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoHorizontalAnchorMixed;
                        var font = dateBox.TextFrame.TextRange.Font;
                        if (font != null)
                        {
                            font.Name = "微软雅黑";
                            font.Size = 14;
                            font.Color.RGB = Color.Gray;
                        }
                    }

                    Console.WriteLine($"  已生成奖状: {data.Name} - {data.Award}");
                }

                Console.WriteLine($"共生成 {presentation.SlideCount} 张奖状幻灯片");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "Certificates.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"奖状演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据合并奖状出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void SlideMasterDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);

                var master = slide.Master;
                if (master != null)
                {
                    Console.WriteLine($"幻灯片母版名称: {master.Name}");
                    Console.WriteLine($"母版宽度: {master.Width} 磅");
                    Console.WriteLine($"母版高度: {master.Height} 磅");

                    var masterShapes = master.Shapes;
                    if (masterShapes != null)
                    {
                        Console.WriteLine($"母版中的形状数量: {masterShapes.Count}");

                        foreach (var shape in masterShapes)
                        {
                            if (shape != null)
                            {
                                Console.WriteLine($"  形状: 名称={shape.Name}, 类型={shape.Type}, 位置=({shape.Left},{shape.Top}), 大小=({shape.Width},{shape.Height})");
                            }
                        }
                    }

                    var customLayouts = master.CustomLayouts;
                    if (customLayouts != null)
                    {
                        Console.WriteLine($"\n母版中的自定义版式数量: {customLayouts.Count}");
                    }

                    var textStyles = master.TextStyles;
                    if (textStyles != null)
                    {
                        Console.WriteLine("母版文本样式可用");
                    }
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "SlideMasterDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"母版操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void PlaceholderDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var titleSlide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (titleSlide != null)
                {
                    var placeholders = titleSlide.Shapes?.Placeholders;
                    if (placeholders != null)
                    {
                        Console.WriteLine($"标题幻灯片中的占位符数量: {placeholders.Count}");

                        for (int i = 1; i <= placeholders.Count; i++)
                        {
                            var placeholder = placeholders[i];
                            if (placeholder != null)
                            {
                                var placeholderFormat = placeholder.PlaceholderFormat;
                                if (placeholderFormat != null)
                                {
                                    Console.WriteLine($"  占位符 {i}: 类型={placeholderFormat.Type}, 名称={placeholder.Name}");
                                }

                                if (placeholder.TextFrame?.TextRange != null)
                                {
                                    string defaultText = placeholder.TextFrame.TextRange.Text ?? "(空)";
                                    Console.WriteLine($"    默认文本: {defaultText}");
                                }
                            }
                        }

                        var titlePlaceholder = placeholders[1];
                        if (titlePlaceholder?.TextFrame?.TextRange != null)
                        {
                            titlePlaceholder.TextFrame.TextRange.Text = "通过占位符设置标题";
                            Console.WriteLine("\n已通过占位符设置标题文本");
                        }

                        var subtitlePlaceholder = placeholders[2];
                        if (subtitlePlaceholder?.TextFrame?.TextRange != null)
                        {
                            subtitlePlaceholder.TextFrame.TextRange.Text = "通过占位符设置副标题";
                            Console.WriteLine("已通过占位符设置副标题文本");
                        }
                    }
                }

                var textSlide = presentation.AddSlide(PpSlideLayout.ppLayoutText);
                if (textSlide != null)
                {
                    var placeholders = textSlide.Shapes?.Placeholders;
                    if (placeholders != null)
                    {
                        Console.WriteLine($"\n文本幻灯片中的占位符数量: {placeholders.Count}");

                        for (int i = 1; i <= placeholders.Count; i++)
                        {
                            var placeholder = placeholders[i];
                            if (placeholder?.PlaceholderFormat != null)
                            {
                                Console.WriteLine($"  占位符 {i}: 类型={placeholder.PlaceholderFormat.Type}");
                            }
                        }

                        var bodyPlaceholder = placeholders[2];
                        if (bodyPlaceholder?.TextFrame?.TextRange != null)
                        {
                            bodyPlaceholder.TextFrame.TextRange.Text = "第一点：通过占位符填充内容\n第二点：保持母版样式一致\n第三点：便于批量修改排版";
                            Console.WriteLine("已通过正文占位符设置内容");
                        }
                    }
                }

                Console.WriteLine("\n提示: 使用 Placeholders 是批量生成复杂排版 PPT 的核心技术");
                Console.WriteLine("  通过占位符填充内容可以保持母版定义的样式和布局");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "PlaceholderDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"占位符操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void BatchReplaceTextDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;

                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutTitle);
                if (slide?.Shapes?.Title?.TextFrame?.TextRange != null)
                    slide.Shapes.Title.TextFrame.TextRange.Text = "{{公司名称}} 年度报告";

                if (slide?.Shapes?.Placeholders != null && slide.Shapes.Placeholders.Count >= 2)
                {
                    var subtitle = slide.Shapes.Placeholders[2];
                    if (subtitle?.TextFrame?.TextRange != null)
                        subtitle.TextFrame.TextRange.Text = "报告人: {{报告人}} | 日期: {{日期}}";
                }

                Console.WriteLine("原始文本已设置（包含占位符标记）");

                int count1 = presentation.ReplaceText("{{公司名称}}", "星辰科技有限公司");
                Console.WriteLine($"替换 {{公司名称}}: {count1} 处");

                int count2 = presentation.ReplaceText("{{报告人}}", "张三");
                Console.WriteLine($"替换 {{报告人}}: {count2} 处");

                int count3 = presentation.ReplaceText("{{日期}}", DateTime.Now.ToString("yyyy年MM月dd日"));
                Console.WriteLine($"替换 {{日期}}: {count3} 处");

                Console.WriteLine($"\n共替换 {count1 + count2 + count3} 处文本");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "BatchReplace.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本替换出错: {ex.Message}");
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
                var generator = new CertificateBatchGenerator();

                var recipients = new[]
                {
                    new CertificateBatchGenerator.Recipient("陈一", "最佳新人奖", "2026年4月"),
                    new CertificateBatchGenerator.Recipient("刘二", "卓越贡献奖", "2026年4月"),
                    new CertificateBatchGenerator.Recipient("孙三", "优秀项目经理", "2026年4月"),
                };

                string filePath = generator.GenerateCertificates(recipients);
                Console.WriteLine($"批量生成奖状完成: {filePath}");
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

    public class CertificateBatchGenerator
    {
        public class Recipient
        {
            public string Name { get; }
            public string Award { get; }
            public string Date { get; }

            public Recipient(string name, string award, string date)
            {
                Name = name;
                Award = award;
                Date = date;
            }
        }

        public string GenerateCertificates(Recipient[] recipients)
        {
            using var app = PowerPointFactory.BlankDocument();
            var presentation = app.ActivePresentation;

            foreach (var recipient in recipients)
            {
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);
                if (slide == null) continue;

                var shapes = slide.Shapes;

                var border = shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 40, 40, 640, 430);
                if (border != null)
                {
                    var fill = border.Fill;
                    fill?.Solid();
                    if (fill?.ForeColor != null)
                        fill.ForeColor.RGB = Color.Ivory;

                    var line = border.Line;
                    if (line?.ForeColor != null)
                    {
                        line.ForeColor.RGB = Color.Goldenrod;
                        line.Weight = 2.5f;
                    }
                }

                var titleBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 80, 560, 60);
                if (titleBox?.TextFrame?.TextRange != null)
                {
                    titleBox.TextFrame.TextRange.Text = "荣誉证书";
                    titleBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                    var font = titleBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "华文行楷";
                        font.Size = 36;
                        font.Bold = true;
                        font.Color.RGB = Color.DarkRed;
                    }
                }

                var nameBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 170, 560, 50);
                if (nameBox?.TextFrame?.TextRange != null)
                {
                    nameBox.TextFrame.TextRange.Text = $"{recipient.Name} 同志";
                    nameBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                    var font = nameBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "微软雅黑";
                        font.Size = 26;
                        font.Bold = true;
                    }
                }

                var awardBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 240, 560, 80);
                if (awardBox?.TextFrame?.TextRange != null)
                {
                    awardBox.TextFrame.TextRange.Text = $"荣获「{recipient.Award}」";
                    awardBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                    var font = awardBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "微软雅黑";
                        font.Size = 20;
                        font.Color.RGB = Color.DarkSlateGray;
                    }
                }

                var dateBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 80, 380, 560, 40);
                if (dateBox?.TextFrame?.TextRange != null)
                {
                    dateBox.TextFrame.TextRange.Text = recipient.Date;
                    dateBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoHorizontalAnchorMixed;
                    var font = dateBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "微软雅黑";
                        font.Size = 14;
                        font.Color.RGB = Color.Gray;
                    }
                }
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);

            string filePath = Path.Combine(tempDirectory, $"Certificates_{Guid.NewGuid():N}.pptx");
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();

            return filePath;
        }
    }
}
