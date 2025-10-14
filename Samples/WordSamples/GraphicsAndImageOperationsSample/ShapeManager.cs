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
    /// 形状管理器类
    /// </summary>
    public class ShapeManager
    {
        private readonly IWordDocument _document;
        private readonly List<IWordShape> _shapes;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public ShapeManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _shapes = new List<IWordShape>();
        }

        /// <summary>
        /// 创建矩形形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateRectangle(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建圆形形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateCircle(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建三角形形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateTriangle(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建箭头形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateArrow(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeCurvedRightArrow, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建星形形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateStar(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeStar32Point, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建云朵形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateCloud(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeCloudCallout, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 创建心形形状
        /// </summary>
        /// <param name="left">左侧位置</param>
        /// <param name="top">顶部位置</param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <param name="text">文本内容</param>
        /// <returns>创建的形状</returns>
        public IWordShape CreateHeart(float left, float top, float width, float height, string text = "")
        {
            var shape = _document.Shapes.AddShape(MsoAutoShapeType.msoShapeHeart, left, top, width, height);
            if (!string.IsNullOrEmpty(text))
            {
                shape.TextFrame.TextRange.Text = text;
            }
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// 批量设置形状填充颜色
        /// </summary>
        /// <param name="color">颜色</param>
        public void SetFillColorForAll(WdColor color)
        {
            foreach (var shape in _shapes)
            {
                if (shape != null && shape.Fill != null)
                {
                    shape.Fill.ForeColor.RGB = (int)color;
                }
            }
        }

        /// <summary>
        /// 批量设置形状边框
        /// </summary>
        /// <param name="color">边框颜色</param>
        /// <param name="weight">边框粗细</param>
        public void SetBorderForAll(WdColor color, float weight)
        {
            foreach (var shape in _shapes)
            {
                if (shape != null && shape.Line != null)
                {
                    shape.Line.ForeColor.RGB = (int)color;
                    shape.Line.Weight = weight;
                }
            }
        }

        /// <summary>
        /// 获取所有形状
        /// </summary>
        /// <returns>形状列表</returns>
        public List<IWordShape> GetAllShapes()
        {
            return new List<IWordShape>(_shapes);
        }

        /// <summary>
        /// 根据索引获取形状
        /// </summary>
        /// <param name="index">索引</param>
        /// <returns>形状对象</returns>
        public IWordShape GetShape(int index)
        {
            if (index >= 0 && index < _shapes.Count)
            {
                return _shapes[index];
            }
            return null;
        }

        /// <summary>
        /// 删除指定形状
        /// </summary>
        /// <param name="shape">要删除的形状</param>
        public void DeleteShape(IWordShape shape)
        {
            if (shape != null && _shapes.Contains(shape))
            {
                shape.Delete();
                _shapes.Remove(shape);
            }
        }

        /// <summary>
        /// 删除所有形状
        /// </summary>
        public void DeleteAllShapes()
        {
            foreach (var shape in _shapes.ToList())
            {
                try
                {
                    shape?.Delete();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"删除形状时出错: {ex.Message}");
                }
            }
            _shapes.Clear();
        }

        /// <summary>
        /// 对齐形状
        /// </summary>
        /// <param name="alignment">对齐方式</param>
        public void AlignShapes(MsoAlignCmd alignment)
        {
            if (_shapes.Count > 1)
            {
                var shapesArray = _shapes.Where(s => s != null).ToArray();
                if (shapesArray.Length > 1)
                {
                    // 将形状数组转换为对象数组
                    object[] shapeObjects = new object[shapesArray.Length];
                    for (int i = 0; i < shapesArray.Length; i++)
                    {
                        shapeObjects[i] = shapesArray[i].Application; // 简化处理
                    }
                    // 注意：实际应用中需要正确传递形状对象
                }
            }
        }

        /// <summary>
        /// 分布形状
        /// </summary>
        /// <param name="distribution">分布方式</param>
        public void DistributeShapes(MsoDistributeCmd distribution)
        {
            if (_shapes.Count > 2)
            {
                var shapesArray = _shapes.Where(s => s != null).ToArray();
                if (shapesArray.Length > 2)
                {
                    // 将形状数组转换为对象数组
                    object[] shapeObjects = new object[shapesArray.Length];
                    for (int i = 0; i < shapesArray.Length; i++)
                    {
                        shapeObjects[i] = shapesArray[i].Application; // 简化处理
                    }
                    // 注意：实际应用中需要正确传递形状对象
                }
            }
        }
    }
}