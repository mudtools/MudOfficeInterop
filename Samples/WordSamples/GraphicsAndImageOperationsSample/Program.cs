//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace GraphicsAndImageOperationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 图形和图片操作示例");

            // 示例1: 图形对象管理
            Console.WriteLine("\n=== 示例1: 图形对象管理 ===");
            GraphicsObjectManagementDemo();

            // 示例2: 图片插入和调整
            Console.WriteLine("\n=== 示例2: 图片插入和调整 ===");
            ImageInsertionAndAdjustmentDemo();

            // 示例3: 形状操作
            Console.WriteLine("\n=== 示例3: 形状操作 ===");
            ShapeOperationsDemo();

            // 示例4: SmartArt图形
            Console.WriteLine("\n=== 示例4: SmartArt图形 ===");
            SmartArtGraphicsDemo();

            // 示例5: 图形效果设置
            Console.WriteLine("\n=== 示例5: 图形效果设置 ===");
            GraphicEffectsDemo();

            // 示例6: 实际应用示例 - 使用辅助类
            Console.WriteLine("\n=== 示例6: 实际应用示例 - 使用辅助类 ===");
            RealWorldGraphicsDemoWithHelpers();

            // 示例7: 完整文档构建示例
            Console.WriteLine("\n=== 示例7: 完整文档构建示例 ===");
            CompleteDocumentBuildDemo();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 图形对象管理示例
        /// </summary>
        static void GraphicsObjectManagementDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 获取内嵌图形集合
                var inlineShapes = document.InlineShapes;

                // 获取浮动图形集合
                var shapes = document.Shapes;

                // 获取图形数量
                int inlineShapeCount = inlineShapes.Count;
                int shapeCount = shapes.Count;

                Console.WriteLine($"初始内嵌图形数量: {inlineShapeCount}");
                Console.WriteLine($"初始浮动图形数量: {shapeCount}");

                // 添加一个内嵌图形进行测试
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                // 由于没有实际图片路径，我们只演示API调用方式
                Console.WriteLine("图形对象管理API演示完成");

                Console.WriteLine("图形对象管理操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"图形对象管理操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 图片插入和调整示例
        /// </summary>
        static void ImageInsertionAndAdjustmentDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加说明文字
                var description = document.Range(document.Content.End - 1, document.Content.End - 1);
                description.Text = "图片插入和调整示例:\n";

                // 尝试插入图片（如果图片存在）
                try
                {
                    // 在文档末尾插入图片
                    var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                    // 由于没有实际图片路径，我们只演示API调用方式
                    // var inlineShape = range.InlineShapes.AddPicture(@"C:\images\example.jpg");
                    Console.WriteLine("图片插入API演示完成");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"图片插入演示: {ex.Message}");
                }

                Console.WriteLine("图片插入和调整操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"图片插入和调整操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 形状操作示例
        /// </summary>
        static void ShapeOperationsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加说明文字
                var description = document.Range(document.Content.End - 1, document.Content.End - 1);
                description.Text = "\n形状操作示例:\n";

                // 添加矩形形状
                var shape1 = document.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    100, 100, 200, 100);

                // 设置形状文本
                shape1.TextFrame.TextRange.Text = "矩形形状";

                // 设置形状填充
                shape1.Fill.ForeColor.RGB = (int)WdColor.wdColorBlue;

                // 设置形状边框
                shape1.Line.ForeColor.RGB = (int)WdColor.wdColorBlack;
                shape1.Line.Weight = 2;

                // 添加圆形形状
                var shape2 = document.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeOval,
                    150, 250, 150, 150);

                shape2.TextFrame.TextRange.Text = "圆形";
                shape2.Fill.ForeColor.RGB = (int)WdColor.wdColorRed;

                // 添加箭头形状
                var shape3 = document.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeCurvedRightArrow,
                    100, 450, 200, 50);

                shape3.TextFrame.TextRange.Text = "箭头";
                shape3.Fill.ForeColor.RGB = (int)WdColor.wdColorGreen;

                Console.WriteLine("形状操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"形状操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// SmartArt图形示例
        /// </summary>
        static void SmartArtGraphicsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加说明文字
                var description = document.Range(document.Content.End - 1, document.Content.End - 1);
                description.Text = "\nSmartArt图形示例:\n";

                // 添加SmartArt图形
                var range = document.Range(document.Content.End - 1, document.Content.End - 1);
                var smartArtShape = document.Shapes.AddSmartArt(
                    null,
                    100, 100, 300, 300);

                // 获取SmartArt对象
                var smartArt = smartArtShape.SmartArt;

                // 添加节点文本
                if (smartArt.AllNodes.Count > 0)
                {
                    smartArt.AllNodes[1].TextFrame2.TextRange.Text = "步骤1";
                }

                if (smartArt.AllNodes.Count > 1)
                {
                    smartArt.AllNodes[2].TextFrame2.TextRange.Text = "步骤2";
                }

                if (smartArt.AllNodes.Count > 2)
                {
                    smartArt.AllNodes[3].TextFrame2.TextRange.Text = "步骤3";
                }


                Console.WriteLine("SmartArt图形操作完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SmartArt图形操作出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 图形效果设置示例
        /// </summary>
        static void GraphicEffectsDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加说明文字
                var description = document.Range(document.Content.End - 1, document.Content.End - 1);
                description.Text = "\n图形效果设置示例:\n";

                // 添加形状
                var shape = document.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    100, 100, 200, 100);

                shape.TextFrame.TextRange.Text = "带效果的形状";

                // 设置阴影效果
                shape.Shadow.Visible = true;
                shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
                shape.Shadow.Blur = 5;
                shape.Shadow.OffsetX = 3;
                shape.Shadow.OffsetY = 3;
                shape.Shadow.ForeColor.RGB = (int)WdColor.wdColorGray50;

                // 设置发光效果
                shape.Glow.Radius = 5;
                shape.Glow.Color.RGB = (int)WdColor.wdColorBlue;

                // 设置柔化边缘效果
                shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
                shape.SoftEdge.Radius = 5;

                // 设置三维格式
                shape.ThreeD.Visible = true;
                shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
                shape.ThreeD.BevelTopInset = 3;
                shape.ThreeD.BevelTopDepth = 2;

                Console.WriteLine("图形效果设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"图形效果设置出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例 - 使用辅助类
        /// </summary>
        static void RealWorldGraphicsDemoWithHelpers()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 在实际应用示例中隐藏Word窗口

                var document = app.ActiveDocument;

                // 创建文档构建器
                var documentBuilder = new GraphicsDocumentBuilder(document);

                // 使用辅助类构建文档
                documentBuilder.BuildDocument("图形操作示例文档(使用辅助类)");

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "GraphicsDemoWithHelpers.docx");
                documentBuilder.SaveDocument(filePath);

                Console.WriteLine($"使用辅助类创建的图形文档已保存: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类创建文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 完整文档构建示例
        /// </summary>
        static void CompleteDocumentBuildDemo()
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                app.Visible = false; // 隐藏Word窗口

                var document = app.ActiveDocument;

                // 创建文档构建器
                var documentBuilder = new GraphicsDocumentBuilder(document);

                // 构建完整文档
                documentBuilder.BuildDocument("完整的图形操作示例文档");

                // 保存文档
                string filePath = Path.Combine(Path.GetTempPath(), "CompleteGraphicsDemo.docx");
                documentBuilder.SaveDocument(filePath);

                Console.WriteLine($"完整图形文档已创建: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建完整文档时出错: {ex.Message}");
            }
        }
    }
}