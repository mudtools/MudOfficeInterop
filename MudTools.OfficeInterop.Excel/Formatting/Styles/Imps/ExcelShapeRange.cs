//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ShapeRange 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.ShapeRange 对象的安全访问和资源管理
/// </summary>
internal class ExcelShapeRange : IExcelShapeRange
{
    /// <summary>
    /// 底层的 COM ShapeRange 对象
    /// </summary>
    private MsExcel.ShapeRange _shapeRange;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelShapeRange 实例
    /// </summary>
    /// <param name="shapeRange">底层的 COM ShapeRange 对象</param>
    internal ExcelShapeRange(MsExcel.ShapeRange shapeRange)
    {
        _shapeRange = shapeRange ?? throw new ArgumentNullException(nameof(shapeRange));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放所有子形状对象
                for (int i = 1; i <= Count; i++)
                {
                    var shape = this[i] as ExcelShape;
                    shape?.Dispose();
                }

                // 释放格式对象
                (_fill as ExcelFillFormat)?.Dispose();
                (_line as ExcelLineFormat)?.Dispose();
                (_textFrame as ExcelTextFrame)?.Dispose();
                (_shadow as ExcelShadowFormat)?.Dispose();
                (_threeD as ExcelThreeDFormat)?.Dispose();
                (_chart as ExcelChart)?.Dispose();

                // 释放底层COM对象
                if (_shapeRange != null)
                    Marshal.ReleaseComObject(_shapeRange);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _shapeRange = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取形状区域中的形状数量
    /// </summary>
    public int Count => _shapeRange?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的形状对象
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    public IExcelShape? this[int index]
    {
        get
        {
            if (_shapeRange == null || index < 1 || index > Count)
                return null;

            try
            {
                var shape = _shapeRange.Item(index);
                return shape != null ? new ExcelShape(shape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    public IExcelShape? this[string name]
    {
        get
        {
            if (_shapeRange == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var shape = _shapeRange.Item(name);
                return shape != null ? new ExcelShape(shape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取或设置形状区域的名称
    /// </summary>
    public string Name
    {
        get => _shapeRange?.Name?.ToString();
        set
        {
            if (_shapeRange != null && value != null)
                _shapeRange.Name = value;
        }
    }

    /// <summary>
    /// 获取形状区域所在的父对象
    /// </summary>
    public object Parent => _shapeRange?.Parent;

    /// <summary>
    /// 获取形状区域的ID
    /// </summary>
    public int ID => _shapeRange?.ID ?? 0;

    #endregion

    #region 位置和大小属性

    /// <summary>
    /// 获取或设置形状区域的左边距
    /// </summary>
    public float Left
    {
        get => _shapeRange?.Left ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Left = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域的顶边距
    /// </summary>
    public float Top
    {
        get => _shapeRange?.Top ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Top = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域的宽度
    /// </summary>
    public float Width
    {
        get => _shapeRange?.Width ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Width = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域的高度
    /// </summary>
    public float Height
    {
        get => _shapeRange?.Height ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Height = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域的旋转角度
    /// </summary>
    public float Rotation
    {
        get => _shapeRange?.Rotation ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Rotation = value;
        }
    }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状区域是否可见
    /// </summary>
    public bool Visible
    {
        get => _shapeRange != null && Convert.ToBoolean(_shapeRange.Visible);
        set
        {
            if (_shapeRange != null)
                _shapeRange.Visible = value ? MsCore.MsoTriState.msoCTrue : MsCore.MsoTriState.msoFalse;
        }
    }
    #endregion

    #region 格式设置

    /// <summary>
    /// 填充格式对象缓存
    /// </summary>
    private IExcelFillFormat _fill;

    /// <summary>
    /// 获取形状区域的填充格式对象
    /// </summary>
    public IExcelFillFormat Fill => _fill ?? (_fill = new ExcelFillFormat(_shapeRange?.Fill));

    /// <summary>
    /// 线条格式对象缓存
    /// </summary>
    private IExcelLineFormat _line;

    /// <summary>
    /// 获取形状区域的线条格式对象
    /// </summary>
    public IExcelLineFormat Line => _line ?? (_line = new ExcelLineFormat(_shapeRange?.Line));

    /// <summary>
    /// 文本框架对象缓存
    /// </summary>
    private IExcelTextFrame _textFrame;

    /// <summary>
    /// 获取形状区域的文本框架对象
    /// </summary>
    public IExcelTextFrame TextFrame => _textFrame ?? (_textFrame = new ExcelTextFrame(_shapeRange?.TextFrame));

    /// <summary>
    /// 阴影格式对象缓存
    /// </summary>
    private IExcelShadowFormat _shadow;

    /// <summary>
    /// 获取形状区域的阴影格式对象
    /// </summary>
    public IExcelShadowFormat Shadow => _shadow ?? (_shadow = new ExcelShadowFormat(_shapeRange?.Shadow));

    /// <summary>
    /// 三维格式对象缓存
    /// </summary>
    private IExcelThreeDFormat _threeD;

    /// <summary>
    /// 获取形状区域的三维格式对象
    /// </summary>
    public IExcelThreeDFormat ThreeD => _threeD ?? (_threeD = new ExcelThreeDFormat(_shapeRange?.ThreeD));

    #endregion

    #region 文本属性

    /// <summary>
    /// 获取或设置形状区域中的文本内容
    /// </summary>
    public string Text
    {
        get => _shapeRange?.TextFrame?.Characters()?.Text?.ToString();
        set
        {
            if (_shapeRange?.TextFrame?.Characters() != null && value != null)
                _shapeRange.TextFrame.Characters().Text = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域中文本的自动调整大小
    /// </summary>
    public bool AutoSize
    {
        get => _shapeRange?.TextFrame != null && Convert.ToBoolean(_shapeRange.TextFrame.AutoSize);
        set
        {
            if (_shapeRange?.TextFrame != null)
                _shapeRange.TextFrame.AutoSize = value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域中文本的水平对齐方式
    /// </summary>
    public XlHAlign HorizontalAlignment
    {
        get => _shapeRange?.TextFrame != null ? (XlHAlign)_shapeRange.TextFrame.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_shapeRange?.TextFrame != null)
                _shapeRange.TextFrame.HorizontalAlignment = (MsExcel.XlHAlign)value;
        }
    }

    /// <summary>
    /// 获取或设置形状区域中文本的垂直对齐方式
    /// </summary>
    public XlVAlign VerticalAlignment
    {
        get => _shapeRange?.TextFrame != null ? (XlVAlign)_shapeRange.TextFrame.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_shapeRange?.TextFrame != null)
                _shapeRange.TextFrame.VerticalAlignment = (MsExcel.XlVAlign)value;
        }
    }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向形状区域添加新的形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape? AddShape(MsoAutoShapeType type, double left, double top, double width, double height)
    {
        if (_shapeRange?.Parent == null) return null;

        try
        {
            // 通过父对象的Shapes集合添加形状
            var parentShapes = _shapeRange.Parent as MsExcel.Shapes;
            if (parentShapes == null) return null;

            var shape = parentShapes.AddShape((MsCore.MsoAutoShapeType)type,
                                            (float)left, (float)top, (float)width, (float)height) as MsExcel.Shape;
            return shape != null ? new ExcelShape(shape) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 向形状区域添加文本框
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的文本框对象</returns>
    public IExcelShape? AddTextbox(MsoTextOrientation orientation, double left, double top, double width, double height)
    {
        if (_shapeRange?.Parent == null) return null;

        try
        {
            var parentShapes = _shapeRange.Parent as MsExcel.Shapes;
            if (parentShapes == null) return null;

            var shape = parentShapes.AddTextbox((MsCore.MsoTextOrientation)orientation,
                                              (float)left, (float)top, (float)width, (float)height) as MsExcel.Shape;
            return shape != null ? new ExcelShape(shape) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 向形状区域添加线条
    /// </summary>
    /// <param name="x1">起点X坐标</param>
    /// <param name="y1">起点Y坐标</param>
    /// <param name="x2">终点X坐标</param>
    /// <param name="y2">终点Y坐标</param>
    /// <returns>新创建的线条对象</returns>
    public IExcelShape AddLine(double x1, double y1, double x2, double y2)
    {
        if (_shapeRange?.Parent == null) return null;

        try
        {
            var parentShapes = _shapeRange.Parent as MsExcel.Shapes;
            if (parentShapes == null) return null;

            var shape = parentShapes.AddLine((float)x1, (float)y1, (float)x2, (float)y2) as MsExcel.Shape;
            return shape != null ? new ExcelShape(shape) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 向形状区域添加图片
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图片对象</returns>
    public IExcelShape AddPicture(string filename, bool linkToFile, bool saveWithDocument,
                                double left, double top, double width, double height)
    {
        if (_shapeRange?.Parent == null || string.IsNullOrEmpty(filename)) return null;

        try
        {
            var parentShapes = _shapeRange.Parent as MsExcel.Shapes;
            if (parentShapes == null) return null;

            var shape = parentShapes.AddPicture(filename,
                linkToFile ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                saveWithDocument ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                (float)left, (float)top, (float)width, (float)height) as MsExcel.Shape;
            return shape != null ? new ExcelShape(shape) : null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 选择和操作

    /// <summary>
    /// 选择形状区域
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    public void Select(bool replace = true)
    {
        _shapeRange?.Select(replace);
    }

    /// <summary>
    /// 删除形状区域中的所有形状
    /// </summary>
    public void Delete()
    {
        _shapeRange?.Delete();
    }

    /// <summary>
    /// 应用自动调整选项
    /// </summary>
    public void Apply()
    {
        _shapeRange?.Apply();
    }

    /// <summary>
    /// 复制形状区域的格式
    /// </summary>
    public void PickUp()
    {
        _shapeRange?.PickUp();
    }
    #endregion

    #region 变换操作

    /// <summary>
    /// 调整形状区域大小
    /// </summary>
    /// <param name="widthScale">宽度缩放比例</param>
    /// <param name="heightScale">高度缩放比例</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小</param>
    public void Scale(double widthScale, double heightScale, bool relativeToOriginalSize = false)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.ScaleWidth((float)widthScale,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                MsExcel.XlScaleType.xlScaleLinear);
            _shapeRange.ScaleHeight((float)heightScale,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                MsExcel.XlScaleType.xlScaleLinear);
        }
        catch
        {
            // 忽略缩放过程中的异常
        }
    }

    /// <summary>
    /// 移动形状区域
    /// </summary>
    /// <param name="leftIncrement">左边距增量</param>
    /// <param name="topIncrement">顶边距增量</param>
    public void Move(double leftIncrement, double topIncrement)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.IncrementLeft((float)leftIncrement);
            _shapeRange.IncrementTop((float)topIncrement);
        }
        catch
        {
            // 忽略移动过程中的异常
        }
    }

    /// <summary>
    /// 旋转形状区域
    /// </summary>
    /// <param name="rotationIncrement">旋转角度增量（度）</param>
    public void Rotate(double rotationIncrement)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.IncrementRotation((float)rotationIncrement);
        }
        catch
        {
            // 忽略旋转过程中的异常
        }
    }


    #endregion

    #region 排列和布局

    public void ScaleHeight(float factor, bool relativeToOriginalSize, MsoScaleFrom? scale = null)
    {
        if (_shapeRange == null) return;

        object scaleObj = Type.Missing;
        if (scale != null)
            scaleObj = (MsCore.MsoScaleFrom)(int)scale;

        _shapeRange.ScaleHeight(Factor: factor,
            relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            Scale: scaleObj);
    }

    public void ScaleWidth(float factor, bool relativeToOriginalSize, MsoScaleFrom? scale = null)
    {
        if (_shapeRange == null) return;
        object scaleObj = Type.Missing;
        if (scale != null)
            scaleObj = (MsCore.MsoScaleFrom)(int)scale;

        _shapeRange.ScaleWidth(Factor: factor,
            relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            Scale: scaleObj);
    }

    /// <summary>
    /// 将形状区域置于最前面
    /// </summary>
    public void BringToFront()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoBringToFront);
    }

    /// <summary>
    /// 将形状区域置于最后面
    /// </summary>
    public void SendToBack()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoSendToBack);
    }

    /// <summary>
    /// 将形状区域向前移动一层
    /// </summary>
    public void BringForward()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoBringForward);
    }

    /// <summary>
    /// 将形状区域向后移动一层
    /// </summary>
    public void SendBackward()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoSendBackward);
    }

    /// <summary>
    /// 对齐形状区域中的形状
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="relativeTo">相对对象</param>
    public void Align(MsoAlignCmd alignment, bool relativeTo = false)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.Align((MsCore.MsoAlignCmd)alignment, relativeTo ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
        }
        catch
        {
            // 忽略对齐过程中的异常
        }
    }

    /// <summary>
    /// 分布形状区域中的形状
    /// </summary>
    /// <param name="distribution">分布方式</param>
    public void Distribute(MsoDistributeCmd distribution)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.Distribute((MsCore.MsoDistributeCmd)distribution, MsCore.MsoTriState.msoFalse);
        }
        catch
        {
            // 忽略分布过程中的异常
        }
    }

    /// <summary>
    /// 统一形状区域中形状的大小
    /// </summary>
    /// <param name="useWidth">是否使用宽度作为标准</param>
    public void SizeToSame(bool useWidth = true)
    {
        if (_shapeRange == null || Count == 0) return;

        try
        {
            // 获取第一个形状的尺寸作为标准
            var firstShape = this[1];
            if (firstShape != null)
            {
                double standardWidth = firstShape.Width;
                double standardHeight = firstShape.Height;

                for (int i = 2; i <= Count; i++)
                {
                    var shape = this[i];
                    if (shape != null)
                    {
                        if (useWidth)
                        {
                            double scale = standardWidth / shape.Width;
                            shape.Scale(scale, scale);
                        }
                        else
                        {
                            double scale = standardHeight / shape.Height;
                            shape.Scale(scale, scale);
                        }
                    }
                }
            }
        }
        catch
        {
            // 忽略统一大小过程中的异常
        }
    }

    #endregion

    #region 组合操作

    /// <summary>
    /// 组合形状区域中的所有形状
    /// </summary>
    /// <returns>组合后的形状对象</returns>
    public IExcelShape Group()
    {
        if (_shapeRange == null) return null;

        try
        {
            var groupedShape = _shapeRange.Group() as MsExcel.Shape;
            return groupedShape != null ? new ExcelShape(groupedShape) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 取消组合形状区域中的组合形状
    /// </summary>
    /// <returns>取消组合后的形状区域</returns>
    public IExcelShapeRange Ungroup()
    {
        if (_shapeRange == null) return null;

        try
        {
            var ungroupedRange = _shapeRange.Ungroup() as MsExcel.ShapeRange;
            return ungroupedRange != null ? new ExcelShapeRange(ungroupedRange) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取形状区域中的所有子形状
    /// </summary>
    /// <returns>子形状数组</returns>
    public IExcelShape[] GetChildShapes()
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            var shape = this[i];
            if (shape != null)
                result.Add(shape);
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取形状区域中的顶级形状
    /// </summary>
    /// <returns>顶级形状数组</returns>
    public IExcelShape[] GetTopLevelShapes()
    {
        // 在ShapeRange中，所有形状都是顶级的
        return GetChildShapes();
    }

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据类型筛选形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <returns>匹配的形状数组</returns>
    public IExcelShape[] FilterByType(MsoShapeType type)
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null && shape.Type == type)
                {
                    result.Add(shape);
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据名称筛选形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的形状数组</returns>
    public IExcelShape[] FilterByName(string name, bool matchCase = false)
    {
        if (_shapeRange == null || string.IsNullOrEmpty(name) || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null && shape.Name != null)
                {
                    bool match = matchCase ?
                        shape.Name.Contains(name) :
                        shape.Name.ToLower().Contains(name.ToLower());

                    if (match)
                        result.Add(shape);
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据位置筛选形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的形状数组</returns>
    public IExcelShape[] FilterByPosition(double left, double top, double tolerance = 10)
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null)
                {
                    double shapeLeft = shape.Left;
                    double shapeTop = shape.Top;

                    if (Math.Abs(shapeLeft - left) <= tolerance && Math.Abs(shapeTop - top) <= tolerance)
                    {
                        result.Add(shape);
                    }
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据大小筛选形状
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的形状数组</returns>
    public IExcelShape[] FilterBySize(double width, double height, double tolerance = 10)
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null)
                {
                    double shapeWidth = shape.Width;
                    double shapeHeight = shape.Height;

                    if (Math.Abs(shapeWidth - width) <= tolerance && Math.Abs(shapeHeight - height) <= tolerance)
                    {
                        result.Add(shape);
                    }
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取可见的形状
    /// </summary>
    /// <returns>可见形状数组</returns>
    public IExcelShape[] GetVisibleShapes()
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null && shape.Visible)
                {
                    result.Add(shape);
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取隐藏的形状
    /// </summary>
    /// <returns>隐藏形状数组</returns>
    public IExcelShape[] GetHiddenShapes()
    {
        if (_shapeRange == null || Count == 0)
            return new IExcelShape[0];

        var result = new List<IExcelShape>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var shape = this[i];
                if (shape != null && !shape.Visible)
                {
                    result.Add(shape);
                }
            }
            catch
            {
                // 忽略单个形状访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 层次结构    
    /// <summary>
    /// 图表对象缓存
    /// </summary>
    private IExcelChart _chart;

    /// <summary>
    /// 获取形状区域所在的图表对象（如果是图表）
    /// </summary>
    public IExcelChart Chart => _chart ?? (_chart = new ExcelChart(_shapeRange?.Chart));

    #endregion


    public IEnumerator<IExcelShape> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}