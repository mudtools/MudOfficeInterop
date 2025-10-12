using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                inlineShape.LockAspectRatio = MsoTriState.msoTrue;
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
        /// 创建SmartArt图形
        /// </summary>
        /// <param name="smartArtType">SmartArt类型</param>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns>创建的SmartArt图形对象</returns>
        public IWordShape CreateSmartArt(MsoSmartArtDefaultConstants smartArtType, float left, float top, float width, float height)
        {
            try
            {
                var range = _document.Range(_document.Content.End - 1, _document.Content.End - 1);
                var smartArtShape = _document.Shapes.AddSmartArt(smartArtType, left, top, width, height);
                return smartArtShape;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建SmartArt时出错: {ex.Message}");
                return null;
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
                shape.Shadow.Visible = MsoTriState.msoTrue;
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
                shape.ThreeD.Visible = MsoTriState.msoTrue;
                shape.ThreeD.BevelTopType = bevelType;
                shape.ThreeD.BevelTopInset = inset;
                shape.ThreeD.BevelTopDepth = depth;
            }
        }
    }
}