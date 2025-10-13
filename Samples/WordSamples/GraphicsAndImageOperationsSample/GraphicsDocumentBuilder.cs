//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace GraphicsAndImageOperationsSample
{
    /// <summary>
    /// 图形文档构建器类
    /// </summary>
    public class GraphicsDocumentBuilder
    {
        private readonly IWordDocument _document;
        private readonly GraphicsHelper _graphicsHelper;
        private readonly ImageProcessor _imageProcessor;
        private readonly ShapeManager _shapeManager;
        private readonly SmartArtHelper _smartArtHelper;
        private readonly GraphicEffectsHelper _effectsHelper;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public GraphicsDocumentBuilder(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _graphicsHelper = new GraphicsHelper(document);
            _imageProcessor = new ImageProcessor(document);
            _shapeManager = new ShapeManager(document);
            _smartArtHelper = new SmartArtHelper(document);
            _effectsHelper = new GraphicEffectsHelper();
        }

        /// <summary>
        /// 添加文档标题
        /// </summary>
        /// <param name="title">标题文本</param>
        public void AddTitle(string title)
        {
            try
            {
                var titleRange = _document.Range();
                titleRange.Text = title + "\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 24;
                titleRange.Font.Bold = true;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加标题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加章节标题
        /// </summary>
        /// <param name="title">章节标题</param>
        /// <param name="level">标题级别</param>
        public void AddSectionTitle(string title, int level = 1)
        {
            try
            {
                var sectionRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                sectionRange.Text = $"\n{title}\n";
                sectionRange.Font.Name = "微软雅黑";
                sectionRange.Font.Bold = true;
                sectionRange.Font.Size = level == 1 ? 18 : 14;
                sectionRange.ParagraphFormat.SpaceAfter = 12;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加章节标题时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加段落文本
        /// </summary>
        /// <param name="text">段落文本</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="alignment">对齐方式</param>
        public void AddParagraph(
            string text,
            string fontName = "宋体",
            float fontSize = 12,
            WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft)
        {
            try
            {
                var paragraphRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                paragraphRange.Text = text + "\n";
                paragraphRange.Font.Name = fontName;
                paragraphRange.Font.Size = fontSize;
                paragraphRange.ParagraphFormat.Alignment = alignment;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加段落时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建图片展示部分
        /// </summary>
        public void CreateImageSection()
        {
            try
            {
                AddSectionTitle("1. 图片操作示例", 2);
                AddParagraph("以下展示了如何在Word文档中插入和调整图片。");

                // 由于没有实际图片路径，我们只添加说明文字
                AddParagraph("[图片插入功能演示 - 需要实际图片路径]", "宋体", 12, WdParagraphAlignment.wdAlignParagraphCenter);
                AddParagraph("图片可以调整大小、亮度、对比度，并设置环绕方式。", "宋体", 10, WdParagraphAlignment.wdAlignParagraphCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建图片展示部分时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建形状展示部分
        /// </summary>
        public void CreateShapesSection()
        {
            try
            {
                AddSectionTitle("2. 形状操作示例", 2);
                AddParagraph("以下展示了如何在Word文档中创建和操作各种形状。");

                // 创建各种形状
                var rectangle = _shapeManager.CreateRectangle(100, 100, 150, 75, "矩形");
                _graphicsHelper.SetShapeFillColor(rectangle, WdColor.wdColorLightBlue);
                _graphicsHelper.SetShapeBorder(rectangle, WdColor.wdColorBlue, 1);

                var circle = _shapeManager.CreateCircle(300, 100, 100, 100, "圆形");
                _graphicsHelper.SetShapeFillColor(circle, WdColor.wdColorLightGreen);
                _graphicsHelper.SetShapeBorder(circle, WdColor.wdColorGreen, 1);

                var triangle = _shapeManager.CreateTriangle(450, 100, 100, 100, "三角形");
                _graphicsHelper.SetShapeFillColor(triangle, WdColor.wdColorLightYellow);
                _graphicsHelper.SetShapeBorder(triangle, WdColor.wdColorOrange, 1);

                var arrow = _shapeManager.CreateArrow(100, 250, 200, 50, "箭头");
                _graphicsHelper.SetShapeFillColor(arrow, WdColor.wdColorDarkRed);
                _graphicsHelper.SetShapeBorder(arrow, WdColor.wdColorRed, 1);

                var star = _shapeManager.CreateStar(350, 250, 100, 100, "星形");
                _graphicsHelper.SetShapeFillColor(star, WdColor.wdColorPink);
                _graphicsHelper.SetShapeBorder(star, WdColor.wdColorViolet, 1);

                var cloud = _shapeManager.CreateCloud(200, 350, 150, 100, "云朵");
                _graphicsHelper.SetShapeFillColor(cloud, WdColor.wdColorLightTurquoise);
                _graphicsHelper.SetShapeBorder(cloud, WdColor.wdColorTeal, 1);

                AddParagraph("\n以上形状展示了基本的形状创建和样式设置功能。", "宋体", 10, WdParagraphAlignment.wdAlignParagraphCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建形状展示部分时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建SmartArt展示部分
        /// </summary>
        public void CreateSmartArtSection()
        {
            try
            {
                AddSectionTitle("3. SmartArt图形示例", 2);
                AddParagraph("以下展示了如何在Word文档中创建和操作SmartArt图形。");

                // 创建列表类型的SmartArt
                var listSmartArt = _smartArtHelper.CreateListSmartArt(100, 100, 400, 300);
                var listItems = new List<string>
                {
                    "项目1: 需求分析",
                    "项目2: 系统设计",
                    "项目3: 编码实现",
                    "项目4: 测试验证",
                    "项目5: 部署上线"
                };
                _smartArtHelper.SetNodeTexts(listSmartArt, listItems);
                _smartArtHelper.SetColorStyle(listSmartArt, 2);
                _smartArtHelper.SetLayoutStyle(listSmartArt, 3);

                AddParagraph("\n以上SmartArt图形展示了项目开发流程。", "宋体", 10, WdParagraphAlignment.wdAlignParagraphCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建SmartArt展示部分时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建图形效果展示部分
        /// </summary>
        public void CreateEffectsSection()
        {
            try
            {
                AddSectionTitle("4. 图形效果示例", 2);
                AddParagraph("以下展示了如何为图形添加各种视觉效果。");

                // 创建基础形状
                var baseShape = _shapeManager.CreateRectangle(100, 100, 150, 75, "基础形状");
                _graphicsHelper.SetShapeFillColor(baseShape, WdColor.wdColorLightBlue);

                // 添加阴影效果
                var shadowShape = _shapeManager.CreateRectangle(300, 100, 150, 75, "阴影效果");
                _graphicsHelper.SetShapeFillColor(shadowShape, WdColor.wdColorLightGreen);
                _effectsHelper.ApplyShadowEffect(shadowShape, GraphicEffectsHelper.ShadowEffectType.OuterShadow, WdColor.wdColorGray50);

                // 添加发光效果
                var glowShape = _shapeManager.CreateRectangle(100, 250, 150, 75, "发光效果");
                _graphicsHelper.SetShapeFillColor(glowShape, WdColor.wdColorLightYellow);
                _effectsHelper.ApplyGlowEffect(glowShape, GraphicEffectsHelper.GlowEffectType.Medium, WdColor.wdColorBlue);

                // 添加三维效果
                var threeDShape = _shapeManager.CreateRectangle(300, 250, 150, 75, "三维效果");
                _graphicsHelper.SetShapeFillColor(threeDShape, WdColor.wdColorDarkRed);
                _effectsHelper.Apply3DFormatEffect(threeDShape, GraphicEffectsHelper.ThreeDEffectType.Simple);

                AddParagraph("\n以上形状展示了不同的图形效果应用。", "宋体", 10, WdParagraphAlignment.wdAlignParagraphCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建图形效果展示部分时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 构建完整图形文档
        /// </summary>
        /// <param name="title">文档标题</param>
        public void BuildDocument(string title = "图形和图片操作示例文档")
        {
            try
            {
                // 添加文档标题
                AddTitle(title);

                // 添加文档说明
                AddParagraph("本文档展示了如何使用MudTools.OfficeInterop.Word库操作Word文档中的图形和图片元素。", "宋体", 12, WdParagraphAlignment.wdAlignParagraphCenter);
                AddParagraph("文档包含图片操作、形状创建、SmartArt图形和视觉效果等多个部分。", "宋体", 12, WdParagraphAlignment.wdAlignParagraphCenter);

                // 创建各个部分
                CreateImageSection();
                CreateShapesSection();
                CreateSmartArtSection();
                CreateEffectsSection();

                // 添加文档结尾
                AddSectionTitle("总结", 2);
                AddParagraph("通过以上示例，我们可以看到MudTools.OfficeInterop.Word库提供了丰富的图形和图片操作功能。");
                AddParagraph("开发者可以利用这些功能创建视觉吸引力强的Word文档，提升文档的专业性和可读性。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"构建文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 保存文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public void SaveDocument(string filePath)
        {
            try
            {
                _document.SaveAs(filePath);
                Console.WriteLine($"文档已保存到: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存文档时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取图形助手
        /// </summary>
        public GraphicsHelper GraphicsHelper => _graphicsHelper;

        /// <summary>
        /// 获取图片处理器
        /// </summary>
        public ImageProcessor ImageProcessor => _imageProcessor;

        /// <summary>
        /// 获取形状管理器
        /// </summary>
        public ShapeManager ShapeManager => _shapeManager;

        /// <summary>
        /// 获取SmartArt助手
        /// </summary>
        public SmartArtHelper SmartArtHelper => _smartArtHelper;

        /// <summary>
        /// 获取图形效果助手
        /// </summary>
        public GraphicEffectsHelper EffectsHelper => _effectsHelper;
    }
}