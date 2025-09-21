//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ShapeRange 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.ShapeRange 对象的安全访问和资源管理
/// </summary>
internal class ExcelShapeRange : IExcelShapeRange
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelShape));
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

        if (disposing && _shapeRange != null)
        {
            try
            {
                // 释放格式对象
                _fill?.Dispose();
                _line?.Dispose();
                _textFrame?.Dispose();
                _shadow?.Dispose();
                _threeD?.Dispose();
                _chart?.Dispose();
                _adjustments?.Dispose();
                _callout?.Dispose();
                _excelConnectorFormat?.Dispose();
                _groupShapes?.Dispose();
                _shapeNodes?.Dispose();
                _effect?.Dispose();
                _pictureFormat?.Dispose();
                _parentGroup?.Dispose();
                _edgeFormat?.Dispose();
                _glow?.Dispose();
                _reflection?.Dispose();

                _fill = null;
                _line = null;
                _textFrame = null;
                _shadow = null;
                _threeD = null;
                _chart = null;
                _adjustments = null; ;
                _callout = null;
                _excelConnectorFormat = null;
                _groupShapes = null;
                _shapeNodes = null;
                _effect = null;
                _pictureFormat = null;
                _edgeFormat = null;
                _parentGroup = null;
                _glow = null;
                _reflection = null;

                // 释放底层COM对象
                if (_shapeRange != null)
                    Marshal.ReleaseComObject(_shapeRange);
            }
            catch (COMException ce)
            {
                log.Error("释放资源时发生异常", ce);
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
    private IExcelFillFormat? _fill;

    /// <summary>
    /// 获取形状区域的填充格式对象
    /// </summary>
    public IExcelFillFormat? Fill
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _fill ??= new ExcelFillFormat(_shapeRange.Fill);
        }
    }

    /// <summary>
    /// 线条格式对象缓存
    /// </summary>
    private IExcelLineFormat? _line;

    /// <summary>
    /// 获取形状区域的线条格式对象
    /// </summary>
    public IExcelLineFormat? Line
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _line ??= new ExcelLineFormat(_shapeRange.Line);
        }
    }

    /// <summary>
    /// 文本框架对象缓存
    /// </summary>
    private IExcelTextFrame? _textFrame;

    /// <summary>
    /// 获取形状区域的文本框架对象
    /// </summary>
    public IExcelTextFrame? TextFrame
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _textFrame ??= new ExcelTextFrame(_shapeRange.TextFrame);
        }
    }

    /// <summary>
    /// 阴影格式对象缓存
    /// </summary>
    private IExcelShadowFormat? _shadow;

    /// <summary>
    /// 获取形状区域的阴影格式对象
    /// </summary>
    public IExcelShadowFormat? Shadow
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _shadow ??= new ExcelShadowFormat(_shapeRange.Shadow);
        }
    }

    /// <summary>
    /// 三维格式对象缓存
    /// </summary>
    private IExcelThreeDFormat? _threeD;

    /// <summary>
    /// 获取形状区域的三维格式对象
    /// </summary>
    public IExcelThreeDFormat? ThreeD
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _threeD ??= new ExcelThreeDFormat(_shapeRange.ThreeD);
        }
    }

    private IExcelAdjustments? _adjustments;

    public IExcelAdjustments? Adjustments
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _adjustments ??= new ExcelAdjustments(_shapeRange.Adjustments);
        }
    }

    private IExcelCalloutFormat? _callout;

    public IExcelCalloutFormat? Callout
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _callout ??= new ExcelCalloutFormat(_shapeRange.Callout);
        }
    }

    private IExcelConnectorFormat? _excelConnectorFormat;

    public IExcelConnectorFormat? ConnectorFormat
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _excelConnectorFormat ??= new ExcelConnectorFormat(_shapeRange.ConnectorFormat);
        }
    }

    private IExcelGroupShapes? _groupShapes;

    public IExcelGroupShapes? GroupItems
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _groupShapes ??= new ExcelGroupShapes(_shapeRange.GroupItems);
        }
    }

    private IExcelShapeNodes? _shapeNodes;

    public IExcelShapeNodes? Nodes
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _shapeNodes ??= new ExcelShapeNodes(_shapeRange.Nodes);
        }
    }

    private IExcelTextEffectFormat? _effect;

    public IExcelTextEffectFormat? TextEffect
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _effect ??= new ExcelTextEffectFormat(_shapeRange.TextEffect);
        }
    }

    private IExcelPictureFormat? _pictureFormat;

    public IExcelPictureFormat? PictureFormat
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _pictureFormat ??= new ExcelPictureFormat(_shapeRange.PictureFormat);
        }
    }

    private IOfficeSoftEdgeFormat? _edgeFormat;

    public IOfficeSoftEdgeFormat? SoftEdge
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _edgeFormat ??= new OfficeSoftEdgeFormat(_shapeRange.SoftEdge);
        }
    }

    private IOfficeGlowFormat? _glow;

    public IOfficeGlowFormat? Glow
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _glow ??= new OfficeGlowFormat(_shapeRange.Glow);
        }
    }

    private IOfficeReflectionFormat? _reflection;
    public IOfficeReflectionFormat? Reflection
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _reflection ??= new OfficeReflectionFormat(_shapeRange.Reflection);
        }
    }

    private IExcelShape? _parentGroup;

    public IExcelShape? ParentGroup
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _parentGroup ??= new ExcelShape(_shapeRange.ParentGroup);
        }
    }

    private IExcelChart? _chart;

    public IExcelChart? Chart
    {
        get
        {
            if (_shapeRange == null)
                return null;
            return _chart ??= new ExcelChart(_shapeRange.Chart);
        }
    }


    public MsoAutoShapeType AutoShapeType
    {
        get => _shapeRange != null ? _shapeRange.AutoShapeType.EnumConvert(MsoAutoShapeType.msoShapeMixed) : MsoAutoShapeType.msoShapeMixed;
        set
        {
            if (_shapeRange != null)
                _shapeRange.AutoShapeType = value.EnumConvert(MsCore.MsoAutoShapeType.msoShapeMixed);
        }
    }

    public MsoShapeStyleIndex ShapeStyle
    {
        get => _shapeRange != null ? _shapeRange.ShapeStyle.EnumConvert(MsoShapeStyleIndex.msoShapeStyleMixed) : MsoShapeStyleIndex.msoShapeStyleMixed;
        set
        {
            if (_shapeRange != null)
                _shapeRange.ShapeStyle = value.EnumConvert(MsCore.MsoShapeStyleIndex.msoShapeStyleMixed);
        }
    }

    public MsoBackgroundStyleIndex BackgroundStyle
    {
        get => _shapeRange != null ? _shapeRange.BackgroundStyle.EnumConvert(MsoBackgroundStyleIndex.msoBackgroundStyleMixed) : MsoBackgroundStyleIndex.msoBackgroundStyleMixed;
        set
        {
            if (_shapeRange != null)
                _shapeRange.BackgroundStyle = value.EnumConvert(MsCore.MsoBackgroundStyleIndex.msoBackgroundStyleMixed);
        }
    }

    public MsoBlackWhiteMode BlackWhiteMode
    {
        get => _shapeRange != null ? _shapeRange.BlackWhiteMode.EnumConvert(MsoBlackWhiteMode.msoBlackWhiteMixed) : MsoBlackWhiteMode.msoBlackWhiteMixed;
        set
        {
            if (_shapeRange != null)
                _shapeRange.BlackWhiteMode = value.EnumConvert(MsCore.MsoBlackWhiteMode.msoBlackWhiteMixed);
        }
    }

    public MsoShapeType Type
    {
        get => _shapeRange != null ? _shapeRange.Type.EnumConvert(MsoShapeType.msoShapeTypeMixed) : MsoShapeType.msoShapeTypeMixed;
    }

    public int ConnectionSiteCount
    {
        get => _shapeRange?.ConnectionSiteCount ?? 0;
    }

    public bool HasChart
    {
        get => _shapeRange != null && _shapeRange.HasChart.ConvertToBool();
    }

    public bool Connector
    {
        get => _shapeRange != null && _shapeRange.Connector.ConvertToBool();
    }

    public bool HorizontalFlip
    {
        get => _shapeRange != null && _shapeRange.HorizontalFlip.ConvertToBool();
    }

    public bool VerticalFlip
    {
        get => _shapeRange != null && _shapeRange.VerticalFlip.ConvertToBool();
    }

    public bool LockAspectRatio
    {
        get => _shapeRange != null && _shapeRange.LockAspectRatio.ConvertToBool();
        set
        {
            if (_shapeRange != null)
                _shapeRange.LockAspectRatio = value.ConvertTriState();
        }
    }

    public int ZOrderPosition
    {
        get => _shapeRange?.ZOrderPosition ?? 0;
    }

    public string Title
    {
        get => _shapeRange != null ? _shapeRange.Title : string.Empty;
        set
        {
            if (_shapeRange?.Title != null && value != null)
                _shapeRange.Title = value;
        }

    }
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
        catch (Exception x)
        {
            log.Error($"添加形状失败: {x.Message}", x);
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
        catch (Exception x)
        {
            log.Error($"添加文本框失败: {x.Message}", x);
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
    public IExcelShape? AddLine(double x1, double y1, double x2, double y2)
    {
        if (_shapeRange?.Parent == null) return null;

        try
        {
            var parentShapes = _shapeRange.Parent as MsExcel.Shapes;
            if (parentShapes == null) return null;

            var shape = parentShapes.AddLine((float)x1, (float)y1, (float)x2, (float)y2) as MsExcel.Shape;
            return shape != null ? new ExcelShape(shape) : null;
        }
        catch (Exception x)
        {
            log.Error($"添加线条失败: {x.Message}", x);
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
    public IExcelShape? AddPicture(string filename, bool linkToFile, bool saveWithDocument,
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
        catch (Exception x)
        {
            log.Error($"添加图片失败: {x.Message}", x);
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
        if (_shapeRange == null)
            return;
        try
        {
            _shapeRange.Select(replace);
        }
        catch (Exception x)
        {
            log.Error($"选择形状区域失败: {x.Message}", x);
        }
    }

    /// <summary>
    /// 删除形状区域中的所有形状
    /// </summary>
    public void Delete()
    {
        if (_shapeRange == null) return;
        try
        {
            _shapeRange.Delete();
        }
        catch (Exception x)
        {
            log.Error($"删除形状区域中的所有形状失败: {x.Message}", x);
        }
    }

    /// <summary>
    /// 应用自动调整选项
    /// </summary>
    public void Apply()
    {
        if (_shapeRange == null) return;
        try
        {
            _shapeRange.Apply();
        }
        catch (Exception x)
        {
            log.Error($"应用自动调整选项失败: {x.Message}", x);
        }
    }

    /// <summary>
    /// 复制形状区域的格式
    /// </summary>
    public void PickUp()
    {
        if (_shapeRange == null) return;
        try
        {
            _shapeRange.PickUp();
        }
        catch (Exception x)
        {
            log.Error($"复制形状区域的格式失败: {x.Message}", x);
        }
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
        catch (Exception x)
        {
            log.Error($"调整形状大小失败: {x.Message}");
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
        catch (Exception x)
        {
            log.Error($"移动形状失败: {x.Message}");
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
        catch (Exception x)
        {
            log.Error($"旋转形状失败: {x.Message}");
        }
    }


    #endregion

    #region 排列和布局

    public void ScaleHeight(float factor, bool relativeToOriginalSize, MsoScaleFrom? scale = null)
    {
        if (_shapeRange == null) return;

        try
        {
            object scaleObj = System.Type.Missing;
            if (scale != null)
                scaleObj = (MsCore.MsoScaleFrom)(int)scale;

            _shapeRange.ScaleHeight(Factor: factor,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                Scale: scaleObj);
        }
        catch (Exception x)
        {
            log.Error($"调整形状高度失败: {x.Message}");
        }

    }

    public void ScaleWidth(float factor, bool relativeToOriginalSize, MsoScaleFrom? scale = null)
    {
        if (_shapeRange == null) return;
        try
        {
            object scaleObj = System.Type.Missing;
            if (scale != null)
                scaleObj = (MsCore.MsoScaleFrom)(int)scale;

            _shapeRange.ScaleWidth(Factor: factor,
                relativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
                Scale: scaleObj);
        }
        catch (Exception x)
        {
            log.Error($"调整形状宽度失败: {x.Message}");
        }
    }

    /// <summary>
    /// 将形状区域置于最前面
    /// </summary>
    public void BringToFront()
    {
        if (_shapeRange == null)
            return;

        try
        {
            _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoBringToFront);
        }
        catch (Exception x)
        {
            log.Error($"将形状置于最顶层失败: {x.Message}");
        }
    }

    /// <summary>
    /// 将形状区域置于最后面
    /// </summary>
    public void SendToBack()
    {
        if (_shapeRange == null)
            return;
        try
        {
            _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoSendToBack);
        }
        catch (Exception x)
        {
            log.Error($"将形状置于最后面失败: {x.Message}");
        }
    }

    /// <summary>
    /// 将形状区域向前移动一层
    /// </summary>
    public void BringForward()
    {
        if (_shapeRange == null)
            return;
        try
        {
            _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoBringForward);
        }
        catch (Exception x)
        {
            log.Error($"将形状向前移动一层失败: {x.Message}");
        }
    }

    /// <summary>
    /// 将形状区域向后移动一层
    /// </summary>
    public void SendBackward()
    {
        if (_shapeRange == null)
            return;
        try
        {
            _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoSendBackward);
        }
        catch (Exception x)
        {
            log.Error($"将形状向后移动一层失败: {x.Message}");
        }
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
        catch (Exception x)
        {
            log.Error($"对齐形状失败: {x.Message}");
        }
    }

    public void SetShapesDefaultProperties()
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.SetShapesDefaultProperties();
        }
        catch (Exception x)
        {
            log.Error($"设置形状默认属性失败: {x.Message}");
        }
    }

    public void IncrementLeft(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.IncrementLeft(Increment);
        }
        catch (Exception x)
        {
            log.Error($"调整形状位置失败: {x.Message}");
        }
    }

    public void IncrementTop(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.IncrementTop(Increment);
        }
        catch (Exception x)
        {
            log.Error($"调整形状位置失败: {x.Message}");
        }
    }

    public void IncrementRotation(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.IncrementRotation(Increment);
        }
        catch (Exception x)
        {
            log.Error($"调整形状旋转失败: {x.Message}");
        }
    }

    public void RerouteConnections()
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.RerouteConnections();
        }
        catch (Exception x)
        {
            log.Error($"重新连接形状失败: {x.Message}");
        }
    }

    public void Flip(MsoFlipCmd FlipCmd)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.Flip(FlipCmd.EnumConvert(MsCore.MsoFlipCmd.msoFlipHorizontal));
        }
        catch (Exception x)
        {
            log.Error($"翻转形状失败: {x.Message}");
        }
    }

    public IExcelShapeRange? Duplicate()
    {
        if (_shapeRange == null) return null;

        try
        {
            var duplicatedShapeRange = _shapeRange.Duplicate();
            return duplicatedShapeRange != null ? new ExcelShapeRange(duplicatedShapeRange) : null;
        }
        catch (Exception x)
        {
            log.Error($"复制形状区域失败: {x.Message}");
            return null;
        }
    }

    public void CanvasCropLeft(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.CanvasCropLeft(Increment);
        }
        catch (Exception x)
        {
            log.Error($"裁剪形状区域失败: {x.Message}");
        }
    }

    public void CanvasCropTop(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.CanvasCropTop(Increment);
        }
        catch (Exception x)
        {
            log.Error($"裁剪形状区域失败: {x.Message}");
        }
    }

    public void CanvasCropRight(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.CanvasCropRight(Increment);
        }
        catch (Exception x)
        {
            log.Error($"裁剪形状区域失败: {x.Message}");
        }
    }

    public void CanvasCropBottom(float Increment)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.CanvasCropBottom(Increment);
        }
        catch (Exception x)
        {
            log.Error($"裁剪形状区域失败: {x.Message}");
        }
    }

    public IExcelShape? Regroup()
    {
        if (_shapeRange == null) return null;

        try
        {
            var regroupedShape = _shapeRange.Regroup();
            return regroupedShape != null ? new ExcelShape(regroupedShape) : null;
        }
        catch (Exception x)
        {
            log.Error($"组合形状失败: {x.Message}");
            return null;
        }
    }

    public void ZOrder(MsoZOrderCmd ZOrderCmd)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.ZOrder(ZOrderCmd.EnumConvert(MsCore.MsoZOrderCmd.msoSendToBack));
        }
        catch (Exception x)
        {
            log.Error($"形状Z轴顺序操作失败: {x.Message}");
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
            _shapeRange.Distribute(distribution.EnumConvert(MsCore.MsoDistributeCmd.msoDistributeHorizontally), MsCore.MsoTriState.msoFalse);
        }
        catch (Exception x)
        {
            log.Error($"分布形状失败: {x.Message}");
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
                            float scale = (float)(standardWidth / shape.Width);
                            shape.Scale(scale, scale);
                        }
                        else
                        {
                            float scale = (float)(standardHeight / shape.Height);
                            shape.Scale(scale, scale);
                        }
                    }
                }
            }
        }
        catch (Exception x)
        {
            log.Error($"统一大小失败: {x.Message}");
        }
    }

    #endregion

    #region 组合操作

    /// <summary>
    /// 组合形状区域中的所有形状
    /// </summary>
    /// <returns>组合后的形状对象</returns>
    public IExcelShape? Group()
    {
        if (_shapeRange == null) return null;

        try
        {
            var groupedShape = _shapeRange.Group();
            return groupedShape != null ? new ExcelShape(groupedShape) : null;
        }
        catch (Exception x)
        {
            log.Error($"组合形状失败: {x.Message}");
            return null;
        }
    }

    /// <summary>
    /// 取消组合形状区域中的组合形状
    /// </summary>
    /// <returns>取消组合后的形状区域</returns>
    public IExcelShapeRange? Ungroup()
    {
        if (_shapeRange == null) return null;

        try
        {
            var ungroupedRange = _shapeRange.Ungroup();
            return ungroupedRange != null ? new ExcelShapeRange(ungroupedRange) : null;
        }
        catch (Exception x)
        {
            log.Error($"取消组合形状失败: {x.Message}");
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
            return [];

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
            catch (Exception x)
            {
                log.Error($"筛选形状失败: {x.Message}");
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
            catch (Exception x)
            {
                log.Error($"按名称查找形状 {name} 时，访问索引为 {i} 的形状发生异常", x);
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
            catch (Exception x)
            {
                log.Error($"按位置查找形状时，访问索引为 {i} 的形状发生异常", x);
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
            catch (Exception x)
            {
                log.Error($"按大小查找形状时，访问索引为 {i} 的形状发生异常", x);
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
            catch (Exception x)
            {
                log.Error($"获取可见形状时，访问索引为 {i} 的形状发生异常", x);
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
            catch (Exception x)
            {
                log.Error($"获取隐藏形状时，访问索引为 {i} 的形状发生异常", x);
            }
        }
        return result.ToArray();
    }
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