
namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// GroupObject（实际为 Shape.Type = msoGroup）COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelGroupObject : IExcelGroupObject
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelGroupObject));
    /// <summary>
    /// 内部持有的原始 COM 对象（实际是 Shape）。
    /// </summary>
    internal MsExcel.GroupObject? _groupObject;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="groupObject">原始的 Shape COM 对象，必须为组合类型（msoGroup），不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 shape 为 null 时抛出。</exception>
    /// <exception cref="ArgumentException">当 Shape 不是组合类型时抛出。</exception>
    internal ExcelGroupObject(MsExcel.GroupObject groupObject)
    {
        if (groupObject == null) throw new ArgumentNullException(nameof(groupObject));

        _groupObject = groupObject;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_groupObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_groupObject);
                }
                catch (Exception ex)
                {
                    log.Error($"释放 GroupObject COM 对象时发生异常: {ex.Message}", ex);
                }
                _groupObject = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet 或 Shapes 集合）。
    /// </summary>
    public object? Parent => _groupObject?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _groupObject?.Application != null
            ? new ExcelApplication(_groupObject.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置组合形状的名称。
    /// </summary>
    public string Name
    {
        get => _groupObject?.Name ?? string.Empty;
        set
        {
            if (_groupObject != null && !string.IsNullOrEmpty(value))
            {
                try
                {
                    _groupObject.Name = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置组合形状名称时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    /// <summary>
    /// 获取组合形状的左上角单元格。
    /// </summary>
    public IExcelRange? TopLeftCell =>
        _groupObject?.TopLeftCell != null
            ? new ExcelRange(_groupObject.TopLeftCell)
            : null;

    /// <summary>
    /// 获取组合形状的右下角单元格。
    /// </summary>
    public IExcelRange? BottomRightCell =>
        _groupObject?.BottomRightCell != null
            ? new ExcelRange(_groupObject.BottomRightCell)
            : null;

    public IExcelShapeRange? ShapeRange
    {
        get
        {
            if (_groupObject == null) return null;
            try
            {
                return _groupObject.ShapeRange is MsExcel.ShapeRange shapeRange ? new ExcelShapeRange(shapeRange) : null;
            }
            catch (Exception ex)
            {
                log.Error($"获取 ShapeRange 时发生异常: {ex.Message}", ex);
                return null;
            }
        }
    }

    public IExcelBorder? Border
    {
        get
        {
            if (_groupObject == null) return null;
            try
            {
                return new ExcelBorder(_groupObject.Border);
            }
            catch (Exception ex)
            {
                log.Error($"获取 Border 时发生异常: {ex.Message}", ex);
                return null;
            }
        }
    }

    public IExcelFont Font
    {
        get
        {
            if (_groupObject == null) return null;
            try
            {
                return new ExcelFont(_groupObject.Font);
            }
            catch (Exception ex)
            {
                log.Error($"获取 Font 时发生异常: {ex.Message}", ex);
                return null;
            }
        }
    }

    public IExcelInterior Interior
    {
        get
        {
            if (_groupObject == null) return null;
            try
            {
                return new ExcelInterior(_groupObject.Interior);
            }
            catch (Exception ex)
            {
                log.Error($"获取 Interior 时发生异常: {ex.Message}", ex);
                return null;
            }
        }
    }



    /// <summary>
    /// 获取组合形状的左边缘位置（单位：磅）。
    /// </summary>
    public double Left
    {
        get => _groupObject != null ? _groupObject.Left : 0f;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.Left = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Left 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public bool? Visible
    {
        get => _groupObject != null ? _groupObject.Visible : null;
        set
        {
            if (_groupObject != null && value != null)
            {
                try
                {
                    _groupObject.Visible = value.Value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Visible 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public bool? AddIndent
    {
        get => _groupObject != null ? _groupObject.AddIndent : null;
        set
        {
            if (_groupObject != null && value != null)
            {
                try
                {
                    _groupObject.AddIndent = value.Value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 AddIndent 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public bool? RoundedCorners
    {
        get => _groupObject != null ? _groupObject.RoundedCorners : null;
        set
        {
            if (_groupObject != null && value != null)
            {
                try
                {
                    _groupObject.RoundedCorners = value.Value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 RoundedCorners 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public bool? Shadow
    {
        get => _groupObject != null ? _groupObject.Shadow : null;
        set
        {
            if (_groupObject != null && value != null)
            {
                try
                {
                    _groupObject.Shadow = value.Value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Shadow 属性时发生异常: {ex.Message}", ex);
                }
            }
        }

    }

    public bool? AutoSize
    {
        get => _groupObject != null ? _groupObject.AutoSize : null;
        set
        {
            if (_groupObject != null && value != null)
            {
                try
                {
                    _groupObject.AutoSize = value.Value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 AutoSize 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public XlArrowHeadLength ArrowHeadLength
    {
        get => _groupObject != null ? _groupObject.ArrowHeadLength.ObjectConvertEnum(XlArrowHeadLength.xlArrowHeadLengthMedium) : XlArrowHeadLength.xlArrowHeadLengthMedium;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.ArrowHeadLength = value.EnumConvert(XlArrowHeadLength.xlArrowHeadLengthMedium);
                }
                catch (Exception ex)
                {
                    log.Error($"设置 ArrowHeadLength 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }
    public XlArrowHeadStyle ArrowHeadStyle
    {
        get => _groupObject != null ? _groupObject.ArrowHeadStyle.ObjectConvertEnum(XlArrowHeadStyle.xlArrowHeadStyleNone) : XlArrowHeadStyle.xlArrowHeadStyleNone;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.ArrowHeadStyle = value.EnumConvert(XlArrowHeadStyle.xlArrowHeadStyleNone);
                }
                catch (Exception ex)
                {
                    log.Error($"设置 ArrowHeadStyle 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public XlArrowHeadWidth ArrowHeadWidth
    {
        get => _groupObject != null ? _groupObject.ArrowHeadWidth.ObjectConvertEnum(XlArrowHeadWidth.xlArrowHeadWidthMedium) : XlArrowHeadWidth.xlArrowHeadWidthMedium;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.ArrowHeadWidth = value.EnumConvert(XlArrowHeadWidth.xlArrowHeadWidthMedium);
                }
                catch (Exception ex)
                {
                    log.Error($"设置 ArrowHeadWidth 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public XlOrientation Orientation
    {
        get => _groupObject != null ? _groupObject.Orientation.ObjectConvertEnum(XlOrientation.xlHorizontal) : XlOrientation.xlHorizontal;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.Orientation = value.EnumConvert(XlOrientation.xlHorizontal);
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Orientation 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    public int ZOrder
    {
        get => _groupObject != null ? _groupObject.ZOrder : 0;
    }

    public int ReadingOrder
    {
        get => _groupObject != null ? _groupObject.ReadingOrder : 0;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.ReadingOrder = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 ReadingOrder 属性时发生异常: {ex.Message}", ex);
                }
            }
        }

    }

    /// <summary>
    /// 获取组合形状的上边缘位置（单位：磅）。
    /// </summary>
    public double Top
    {
        get => _groupObject != null ? _groupObject.Top : 0f;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.Top = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Top 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    /// <summary>
    /// 获取或设置组合形状的宽度（单位：磅）。
    /// </summary>
    public double Width
    {
        get => _groupObject != null ? _groupObject.Width : 0f;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.Width = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Width 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    /// <summary>
    /// 获取或设置组合形状的高度（单位：磅）。
    /// </summary>
    public double Height
    {
        get => _groupObject != null ? _groupObject.Height : 0f;
        set
        {
            if (_groupObject != null)
            {
                try
                {
                    _groupObject.Height = value;
                }
                catch (Exception ex)
                {
                    log.Error($"设置 Height 属性时发生异常: {ex.Message}", ex);
                }
            }
        }
    }

    /// <summary>
    /// 取消组合，将组合形状拆分为独立的子形状。
    /// 返回新创建的 Shapes 集合（包含所有拆分后的子项）。
    /// </summary>
    /// <returns>拆分后的子形状集合。</returns>
    public IExcelShapes? Ungroup()
    {
        if (_groupObject == null) return null;
        try
        {
            return _groupObject.Ungroup() is MsExcel.Shapes shapes ? new ExcelShapes(shapes) : null;
        }
        catch (Exception ex)
        {
            log.Error($"取消组合形状时发生异常: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// 将组合形状置于所有其他形状的顶层。
    /// </summary>
    public void BringToFront()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject?.BringToFront();
        }
        catch (Exception ex)
        {
            log.Error($"将组合形状置于顶层时发生异常: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 将组合形状置于所有其他形状的底层。
    /// </summary>
    public void SendToBack()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject?.SendToBack();
        }
        catch (Exception ex)
        {
            log.Error($"将组合形状置于底层时发生异常: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 删除此组合形状（及其所有子项）。
    /// </summary>
    public void Delete()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除组合形状时发生异常: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 选择此组合形状（激活并选中）。
    /// </summary>
    public void Select()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject.Select();
        }
        catch (Exception ex)
        {
            log.Error($"选择组合形状时发生异常: {ex.Message}", ex);
        }
    }

    public void Copy()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject.Copy();
        }
        catch (Exception ex)
        {
            log.Error($"复制组合形状时发生异常: {ex.Message}", ex);
        }
    }

    public void Cut()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject.Cut();
        }
        catch (Exception ex)
        {
            log.Error($"剪切组合形状时发生异常: {ex.Message}", ex);
        }
    }

    public void Duplicate()
    {
        if (_groupObject == null) return;
        try
        {
            _groupObject.Duplicate();
        }
        catch (Exception ex)
        {
            log.Error($"复制组合形状时发生异常: {ex.Message}", ex);
        }
    }

    public void CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter,
                           XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture)
    {
        if (_groupObject == null)
            return;
        try
        {
            _groupObject.CopyPicture(Appearance.EnumConvert(MsExcel.XlPictureAppearance.xlPrinter),
            Format.EnumConvert(MsExcel.XlCopyPictureFormat.xlPicture));
        }
        catch (Exception ex)
        {
            log.Error($"将组合形状复制为图片时发生异常: {ex.Message}", ex);
        }
    }
}