//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.PowerPoint;
using System.Drawing;

namespace ShapeOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.PowerPoint - Shape (形状) 操作示例");

            Console.WriteLine("\n=== 示例1: 文本框操作 ===");
            TextboxDemo();

            Console.WriteLine("\n=== 示例2: 文本格式设置 ===");
            TextFormattingDemo();

            Console.WriteLine("\n=== 示例3: 图像操作 ===");
            PictureDemo();

            Console.WriteLine("\n=== 示例4: 表格操作 ===");
            TableDemo();

            Console.WriteLine("\n=== 示例5: 图表操作 ===");
            ChartDemo();

            Console.WriteLine("\n=== 示例6: 自选图形与连接线 ===");
            AutoShapeDemo();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void TextboxDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var textbox1 = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 600, 40);
                if (textbox1?.TextFrame?.TextRange != null)
                {
                    textbox1.TextFrame.TextRange.Text = "这是一个水平文本框";
                    Console.WriteLine("已添加水平文本框");
                }

                var textbox2 = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 120, 600, 100);
                if (textbox2?.TextFrame?.TextRange != null)
                {
                    textbox2.TextFrame.TextRange.Text = "第一行文本\n第二行文本\n第三行文本";
                    Console.WriteLine("已添加多行文本框");
                }

                var textbox3 = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 260, 600, 40);
                if (textbox3?.TextFrame?.TextRange != null)
                {
                    textbox3.TextFrame.TextRange.Text = "使用 InsertAfter 追加文本 - ";
                    textbox3.TextFrame.TextRange.InsertAfter("追加的内容");
                    textbox3.TextFrame.TextRange.InsertAfter(" | 继续追加");
                    Console.WriteLine("已添加使用 InsertAfter 的文本框");
                }

                Console.WriteLine($"当前幻灯片共有 {shapes.Count} 个形状");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "TextboxDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本框操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void TextFormattingDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var titleBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 30, 600, 50);
                if (titleBox?.TextFrame?.TextRange != null)
                {
                    titleBox.TextFrame.TextRange.Text = "文本格式设置演示";
                    var font = titleBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "微软雅黑";
                        font.Size = 28;
                        font.Bold = true;
                        font.Color.RGB = Color.DarkBlue;
                    }
                    Console.WriteLine("已设置标题: 微软雅黑 28号 加粗 深蓝色");
                }

                var normalBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 100, 600, 40);
                if (normalBox?.TextFrame?.TextRange != null)
                {
                    normalBox.TextFrame.TextRange.Text = "这是普通文本 - 宋体 18号";
                    var font = normalBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "宋体";
                        font.Size = 18;
                        font.Color.RGB = Color.Black;
                    }
                }

                var italicBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 160, 600, 40);
                if (italicBox?.TextFrame?.TextRange != null)
                {
                    italicBox.TextFrame.TextRange.Text = "这是斜体文本 - 楷体 20号 斜体 红色";
                    var font = italicBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "楷体";
                        font.Size = 20;
                        font.Italic = true;
                        font.Color.RGB = Color.Red;
                    }
                }

                var underlineBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 220, 600, 40);
                if (underlineBox?.TextFrame?.TextRange != null)
                {
                    underlineBox.TextFrame.TextRange.Text = "这是带下划线的文本";
                    var font = underlineBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "黑体";
                        font.Size = 18;
                        font.Underline = true;
                        font.Color.RGB = Color.Green;
                    }
                }

                var shadowBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 280, 600, 40);
                if (shadowBox?.TextFrame?.TextRange != null)
                {
                    shadowBox.TextFrame.TextRange.Text = "这是带阴影的文本";
                    var font = shadowBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Name = "微软雅黑";
                        font.Size = 22;
                        font.Shadow = true;
                        font.Color.RGB = Color.Purple;
                    }
                }

                var textFrameBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 340, 300, 80);
                if (textFrameBox?.TextFrame != null)
                {
                    textFrameBox.TextFrame.TextRange.Text = "文本框对齐与自动调整";
                    textFrameBox.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                    textFrameBox.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    textFrameBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    textFrameBox.TextFrame.WordWrap = true;
                    Console.WriteLine("已设置文本框居中对齐和自动调整大小");
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "TextFormatting.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文本格式设置出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void PictureDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                string tempDirectory = GetTempDirectory();
                string imagePath = Path.Combine(tempDirectory, "sample_image.png");

                CreateSampleImage(imagePath);

                var picture = shapes.AddPicture(imagePath, false, true, 50, 50, 300, 200);
                if (picture != null)
                {
                    Console.WriteLine($"已插入图片: 位置(Left={picture.Left}, Top={picture.Top}), 大小(Width={picture.Width}, Height={picture.Height})");
                }

                var picture2 = shapes.AddPicture(imagePath, false, true, 400, 50, 200, 133);
                if (picture2 != null)
                {
                    Console.WriteLine($"已插入缩放图片: 位置(Left={picture2.Left}, Top={picture2.Top}), 大小(Width={picture2.Width}, Height={picture2.Height})");
                }

                Console.WriteLine("\n提示: PPT内部单位为磅(Point)");
                Console.WriteLine("  1厘米 ≈ 28.35磅");
                Console.WriteLine("  1英寸 = 72磅");
                Console.WriteLine("  标准幻灯片(16:9)约为 960×540 磅");

                string filePath = Path.Combine(tempDirectory, "PictureDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"图像操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void TableDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var tableShape = shapes.AddTable(4, 3, 50, 50, 600, 200);
                if (tableShape != null && tableShape.HasTable)
                {
                    var table = tableShape.Table;
                    Console.WriteLine($"已创建 {4}×{3} 表格");

                    string[,] data = {
                        { "姓名", "部门", "业绩" },
                        { "张三", "销售部", "120万" },
                        { "李四", "技术部", "95万" },
                        { "王五", "市场部", "88万" }
                    };

                    for (int row = 1; row <= 4; row++)
                    {
                        for (int col = 1; col <= 3; col++)
                        {
                            var cell = table.Cell(row, col);
                            if (cell?.Shape?.TextFrame?.TextRange != null)
                            {
                                cell.Shape.TextFrame.TextRange.Text = data[row - 1, col - 1];

                                if (row == 1)
                                {
                                    var font = cell.Shape.TextFrame.TextRange.Font;
                                    if (font != null)
                                    {
                                        font.Bold = true;
                                        font.Size = 14;
                                        font.Color.RGB = Color.White;
                                    }
                                    var fill = cell.Shape.Fill;
                                    fill?.Solid();
                                    if (fill?.ForeColor != null)
                                        fill.ForeColor.RGB = Color.SteelBlue;
                                }
                            }
                        }
                    }

                    Console.WriteLine("表格数据已填充，表头已设置格式");

                    table.FirstRow = true;
                    table.HorizBanding = true;
                    Console.WriteLine("已启用首行特殊格式和水平条纹");
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "TableDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"表格操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void ChartDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var titleBox = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 20, 600, 40);
                if (titleBox?.TextFrame?.TextRange != null)
                {
                    titleBox.TextFrame.TextRange.Text = "图表操作示例";
                    var font = titleBox.TextFrame.TextRange.Font;
                    if (font != null)
                    {
                        font.Size = 24;
                        font.Bold = true;
                    }
                }

                var chartShape = shapes.AddChart(XlChartType.xlColumnClustered, 50, 80, 600, 350);
                if (chartShape != null)
                {
                    Console.WriteLine("已添加柱状图 (xlColumnClustered)");
                    Console.WriteLine($"图表形状位置: Left={chartShape.Left}, Top={chartShape.Top}");
                    Console.WriteLine($"图表形状大小: Width={chartShape.Width}, Height={chartShape.Height}");
                    Console.WriteLine("提示: 图表数据需要通过 Chart 对象的 Worksheet 进行编辑");
                }

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "ChartDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"图表操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void AutoShapeDemo()
        {
            IPowerPointApplication? app = null;
            try
            {
                app = PowerPointFactory.BlankDocument();
                app.Visible = true;

                var presentation = app.ActivePresentation;
                var slide = presentation.AddSlide(PpSlideLayout.ppLayoutBlank);

                var shapes = slide.Shapes;

                var rect = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 50, 200, 100);
                if (rect != null)
                {
                    rect.Name = "矩形1";
                    if (rect.TextFrame?.TextRange != null)
                        rect.TextFrame.TextRange.Text = "矩形";
                    Console.WriteLine("已添加矩形");
                }

                var roundedRect = shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 300, 50, 200, 100);
                if (roundedRect != null)
                {
                    if (roundedRect.TextFrame?.TextRange != null)
                        roundedRect.TextFrame.TextRange.Text = "圆角矩形";
                    Console.WriteLine("已添加圆角矩形");
                }

                var oval = shapes.AddShape(MsoAutoShapeType.msoShapeOval, 50, 200, 150, 150);
                if (oval != null)
                {
                    if (oval.TextFrame?.TextRange != null)
                        oval.TextFrame.TextRange.Text = "椭圆";
                    Console.WriteLine("已添加椭圆");
                }

                var arrow = shapes.AddShape(MsoAutoShapeType.msoShapeChevron, 250, 230, 200, 60);
                if (arrow != null)
                {
                    if (arrow.TextFrame?.TextRange != null)
                        arrow.TextFrame.TextRange.Text = "箭头";
                    Console.WriteLine("已添加箭头");
                }

                var connector = shapes.AddConnector(MsoConnectorType.msoConnectorStraight, 50, 370, 500, 370);
                if (connector != null)
                {
                    Console.WriteLine("已添加直线连接线");
                }

                var line = shapes.AddLine(50, 400, 500, 400);
                if (line != null)
                {
                    Console.WriteLine("已添加直线");
                }

                var wordArt = shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect1, "艺术字示例", "微软雅黑", 36, true, false, 50, 440);
                if (wordArt != null)
                {
                    Console.WriteLine("已添加艺术字");
                }

                Console.WriteLine($"\n当前幻灯片共有 {shapes.Count} 个形状");

                string tempDirectory = GetTempDirectory();
                string filePath = Path.Combine(tempDirectory, "AutoShapeDemo.pptx");
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Console.WriteLine($"演示文稿已保存到: {filePath}");

                presentation.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自选图形操作出错: {ex.Message}");
            }
            finally
            {
                app?.Dispose();
            }
        }

        static void CreateSampleImage(string path)
        {
            using var bmp = new System.Drawing.Bitmap(400, 267);
            using var g = System.Drawing.Graphics.FromImage(bmp);
            g.Clear(Color.LightSkyBlue);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            using var pen = new Pen(Color.DarkBlue, 3);
            g.DrawRectangle(pen, 20, 20, 360, 227);

            using var font = new Font("微软雅黑", 24, FontStyle.Bold);
            using var brush = new SolidBrush(Color.DarkBlue);
            var textSize = g.MeasureString("示例图片", font);
            g.DrawString("示例图片", font, brush, (400 - textSize.Width) / 2, (267 - textSize.Height) / 2);

            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }

        static string GetTempDirectory()
        {
            string tempDirectory = Path.Combine(AppContext.BaseDirectory, "Output\\PowerPointSamples");
            if (!Directory.Exists(tempDirectory))
                Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }
}
