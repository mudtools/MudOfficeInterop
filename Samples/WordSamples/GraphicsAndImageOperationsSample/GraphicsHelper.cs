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
    /// <summary>
    /// 图形操作辅助类
    /// </summary>
    public class GraphicsHelper
    {
        private readonly IWordDocument _document;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public GraphicsHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 插入图片到指定位置
        /// </summary>
        /// <param name="imagePath">图片路径</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        /// <param name="position">插入位置</param>
        /// <returns>插入的内嵌图形对象</returns>
        public IWordInlineShape InsertImage(string imagePath, float width, float height, IWordRange position)
        {
            try
            {
                var inlineShape = position.InlineShapes.AddPicture(imagePath);
                inlineShape.Width = width;
                inlineShape.Height = height;
                inlineShape.LockAspectRatio = true;
                return inlineShape;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入图片时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 创建基本形状
        /// </summary>
        /// <param name="shapeType">形状类型</param>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">形状文本</param>
        /// <returns>创建的形状对象</returns>
        public IWordShape CreateShape(MsoAutoShapeType shapeType, float left, float top, float width, float height, string text)
        {
            try
            {
                var shape = _document.Shapes.AddShape(shapeType, left, top, width, height);
                shape.TextFrame.TextRange.Text = text;
                return shape;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建形状时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 设置形状填充颜色
        /// </summary>
        /// <param name="shape">形状对象</param>
        /// <param name="color">颜色</param>
        public void SetShapeFillColor(IWordShape shape, WdColor color)
        {
            if (shape != null)
            {
                shape.Fill.ForeColor.RGB = (int)color;
            }
        }

        /// <summary>
        /// 设置形状边框
        /// </summary>
        /// <param name="shape">形状对象</param>
        /// <param name="color">边框颜色</param>
        /// <param name="weight">边框粗细</param>
        public void SetShapeBorder(IWordShape shape, WdColor color, float weight)
        {
            if (shape != null)
            {
                shape.Line.ForeColor.RGB = (int)color;
                shape.Line.Weight = weight;
            }
        }


        /// <summary>
        /// 为图形添加阴影效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="offsetX">阴影X偏移</param>
        /// <param name="offsetY">阴影Y偏移</param>
        /// <param name="blur">阴影模糊度</param>
        /// <param name="color">阴影颜色</param>
        public void AddShadowEffect(IWordShape shape, float offsetX, float offsetY, float blur, WdColor color)
        {
            if (shape != null)
            {
                shape.Shadow.Visible = true;
                shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
                shape.Shadow.Blur = blur;
                shape.Shadow.OffsetX = offsetX;
                shape.Shadow.OffsetY = offsetY;
                shape.Shadow.ForeColor.RGB = (int)color;
            }
        }

        /// <summary>
        /// 为图形添加发光效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="radius">发光半径</param>
        /// <param name="color">发光颜色</param>
        public void AddGlowEffect(IWordShape shape, float radius, WdColor color)
        {
            if (shape != null)
            {
                shape.Glow.Radius = radius;
                shape.Glow.Color.RGB = (int)color;
            }
        }

        /// <summary>
        /// 为图形添加柔化边缘效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="radius">柔化半径</param>
        public void AddSoftEdgeEffect(IWordShape shape, float radius)
        {
            if (shape != null)
            {
                shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
                shape.SoftEdge.Radius = radius;
            }
        }

        /// <summary>
        /// 为图形添加三维效果
        /// </summary>
        /// <param name="shape">图形对象</param>
        /// <param name="bevelType">斜面类型</param>
        /// <param name="inset">斜面内缩</param>
        /// <param name="depth">斜面深度</param>
        public void Add3DEffect(IWordShape shape, MsoBevelType bevelType, float inset, float depth)
        {
            if (shape != null)
            {
                shape.ThreeD.Visible = true;
                shape.ThreeD.BevelTopType = bevelType;
                shape.ThreeD.BevelTopInset = inset;
                shape.ThreeD.BevelTopDepth = depth;
            }
        }
    }
}