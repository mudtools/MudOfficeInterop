//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Style 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Style 对象的安全访问和资源管理
/// </summary>
internal class ExcelStyle : IExcelStyle
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelStyle));
    /// <summary>
    /// 底层的 COM Style 对象
    /// </summary>
    internal MsExcel.Style? _style;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelStyle 实例
    /// </summary>
    /// <param name="style">底层的 COM Style 对象</param>
    internal ExcelStyle(MsExcel.Style style)
    {
        _style = style ?? throw new ArgumentNullException(nameof(style));
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
                // 释放子COM组件
                (_font as ExcelFont)?.Dispose();
                (_borders as ExcelBorders)?.Dispose();
                (_interior as ExcelInterior)?.Dispose();

                // 释放底层COM对象
                if (_style != null)
                    Marshal.ReleaseComObject(_style);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _style = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    public string NameLocal => _style != null ? _style.NameLocal : string.Empty;

    /// <summary>
    /// 获取样式名称
    /// </summary>
    public string Name => _style != null ? _style.Name : string.Empty;


    /// <summary>
    /// 获取样式是否为内置样式
    /// </summary>
    public bool BuiltIn => _style != null && _style.BuiltIn;

    /// <summary>
    /// 获取样式所在的父对象
    /// </summary>
    public object? Parent => _style?.Parent;

    /// <summary>
    /// 获取样式所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            MsExcel.Application? application = _style?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }
    #endregion

    #region 格式属性

    public bool IncludeNumber
    {
        get => _style != null && _style.IncludeNumber;
        set
        {
            if (_style != null)
                _style.IncludeNumber = value;
        }
    }

    public bool IncludeFont
    {
        get => _style != null && _style.IncludeFont;
        set
        {
            if (_style != null)
                _style.IncludeFont = value;
        }
    }


    public bool IncludeAlignment
    {
        get => _style != null && _style.IncludeAlignment;
        set
        {
            if (_style != null)
                _style.IncludeAlignment = value;
        }
    }

    public bool AddIndent
    {
        get => _style != null && _style.AddIndent;
        set
        {
            if (_style != null)
                _style.AddIndent = value;
        }
    }



    /// <summary>
    /// 字体对象缓存
    /// </summary>
    private IExcelFont _font;

    /// <summary>
    /// 获取样式的字体对象
    /// </summary>
    public IExcelFont? Font
    {
        get
        {
            if (_style == null)
                return null;
            _font ??= new ExcelFont(_style.Font);
            return _font;
        }
    }

    /// <summary>
    /// 边框对象缓存
    /// </summary>
    private IExcelBorders _borders;

    /// <summary>
    /// 获取样式的边框对象
    /// </summary>
    public IExcelBorders? Borders
    {
        get
        {
            if (_style == null)
                return null;
            _borders ??= new ExcelBorders(_style.Borders);
            return _borders;
        }
    }

    /// <summary>
    /// 内部格式对象缓存
    /// </summary>
    private IExcelInterior _interior;

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior? Interior
    {
        get
        {
            if (_style == null)
                return null;
            _interior ??= new ExcelInterior(_style.Interior);
            return _interior;
        }
    }

    /// <summary>
    /// 获取或设置样式的数字格式
    /// </summary>
    public string NumberFormat
    {
        get => _style != null ? _style.NumberFormat : string.Empty;
        set
        {
            if (_style != null && value != null)
                _style.NumberFormat = value;
        }
    }

    public string NumberFormatLocal
    {
        get => _style != null ? _style.NumberFormatLocal : string.Empty;
        set
        {
            if (_style != null && value != null)
                _style.NumberFormatLocal = value;
        }
    }

    /// <summary>
    /// 获取或设置样式的水平对齐方式
    /// </summary>
    public XlHAlign HorizontalAlignment
    {
        get => _style != null ? _style.HorizontalAlignment.EnumConvert(XlHAlign.xlHAlignGeneral) : XlHAlign.xlHAlignGeneral;
        set
        {
            if (_style != null)
                _style.HorizontalAlignment = value.EnumConvert(MsExcel.XlHAlign.xlHAlignGeneral);
        }
    }

    /// <summary>
    /// 获取或设置样式的垂直对齐方式
    /// </summary>
    public XlVAlign VerticalAlignment
    {
        get => _style != null ? _style.VerticalAlignment.EnumConvert(XlVAlign.xlVAlignCenter) : XlVAlign.xlVAlignCenter;
        set
        {
            if (_style != null)
                _style.VerticalAlignment = value.EnumConvert(MsExcel.XlVAlign.xlVAlignCenter);
        }
    }

    /// <summary>
    /// 获取或设置样式是否自动换行
    /// </summary>
    public bool WrapText
    {
        get => _style != null && Convert.ToBoolean(_style.WrapText);
        set
        {
            if (_style != null)
                _style.WrapText = value;
        }
    }

    /// <summary>
    /// 获取或设置样式的缩进级别
    /// </summary>
    public int IndentLevel
    {
        get => _style != null ? Convert.ToInt32(_style.IndentLevel) : 0;
        set
        {
            if (_style != null)
                _style.IndentLevel = value;
        }
    }

    /// <summary>
    /// 获取或设置样式的阅读顺序
    /// </summary>
    public int ReadingOrder
    {
        get => _style != null ? Convert.ToInt32(_style.ReadingOrder) : 0;
        set
        {
            if (_style != null)
                _style.ReadingOrder = value;
        }
    }

    /// <summary>
    /// 获取或设置样式的旋转角度
    /// </summary>
    public XlOrientation Orientation
    {
        get => _style != null ? _style.Orientation.EnumConvert(XlOrientation.xlHorizontal) : XlOrientation.xlHorizontal;
        set
        {
            if (_style != null)
                _style.Orientation = value.EnumConvert(MsExcel.XlOrientation.xlHorizontal);
        }
    }

    /// <summary>
    /// 获取或设置样式是否收缩适应
    /// </summary>
    public bool ShrinkToFit
    {
        get => _style != null && _style.ShrinkToFit;
        set
        {
            if (_style != null)
                _style.ShrinkToFit = value;
        }
    }

    /// <summary>
    /// 获取或设置样式是否合并单元格
    /// </summary>
    public bool MergeCells
    {
        get => _style != null && _style.MergeCells.ConvertToBool();
        set
        {
            if (_style != null)
                _style.MergeCells = value;
        }
    }

    /// <summary>
    /// 获取或设置样式是否锁定
    /// </summary>
    public bool Locked
    {
        get => _style != null && _style.Locked;
        set
        {
            if (_style != null)
                _style.Locked = value;
        }
    }

    /// <summary>
    /// 获取或设置样式是否隐藏公式
    /// </summary>
    public bool FormulaHidden
    {
        get => _style != null && _style.FormulaHidden;
        set
        {
            if (_style != null)
                _style.FormulaHidden = value;
        }
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除样式
    /// </summary>
    public void Delete()
    {
        if (_style == null)
            return;
        try
        {
            _style.Delete();
        }
        catch (Exception ex)
        {
            log.Error("删除样式对象时发生异常", ex);
        }
    }

    /// <summary>
    /// 复制样式
    /// </summary>
    /// <param name="newName">新样式名称</param>
    /// <returns>复制的样式对象</returns>
    public IExcelStyle? Copy(string newName)
    {
        if (_style?.Parent == null || string.IsNullOrEmpty(newName))
            return null;

        try
        {
            var parentStyles = _style.Parent as MsExcel.Styles;
            if (parentStyles != null)
            {
                var newStyle = parentStyles.Add(newName) as MsExcel.Style;
                if (newStyle != null)
                {
                    var excelStyle = new ExcelStyle(newStyle);

                    // 复制样式属性
                    excelStyle.Font.Name = Font.Name;
                    excelStyle.Font.Size = Font.Size;
                    excelStyle.Font.Bold = Font.Bold;
                    excelStyle.Font.Italic = Font.Italic;
                    excelStyle.Font.Color = Font.Color;
                    excelStyle.Font.Underline = Font.Underline;
                    excelStyle.NumberFormat = NumberFormat;
                    excelStyle.HorizontalAlignment = HorizontalAlignment;
                    excelStyle.VerticalAlignment = VerticalAlignment;
                    excelStyle.WrapText = WrapText;
                    excelStyle.IndentLevel = IndentLevel;
                    excelStyle.Orientation = Orientation;
                    excelStyle.ShrinkToFit = ShrinkToFit;
                    excelStyle.MergeCells = MergeCells;
                    excelStyle.Locked = Locked;
                    excelStyle.FormulaHidden = FormulaHidden;

                    return excelStyle;
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            log.Error($"复制样式 {Name} 到 {newName} 时发生异常", ex);
            return null;
        }
    }

    /// <summary>
    /// 重命名样式
    /// </summary>
    /// <param name="newName">新样式名称</param>
    public void Rename(string newName)
    {
        if (_style == null || string.IsNullOrEmpty(newName))
            return;

        try
        {
            // 通过复制并删除实现重命名
            var parentStyles = _style.Parent as MsExcel.Styles;
            if (parentStyles != null)
            {
                var newStyle = Copy(newName);
                if (newStyle != null)
                {
                    _style.Delete();
                    _style = (newStyle as ExcelStyle)?._style;
                }
            }
        }
        catch (Exception ex)
        {
            log.Error($"重命名样式 {Name} 为 {newName} 时发生异常", ex);
        }
    }

    /// <summary>
    /// 应用样式到指定区域
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <param name="includeFont">是否包含字体</param>
    /// <param name="includeBorder">是否包含边框</param>
    /// <param name="includeFill">是否包含填充</param>
    public void ApplyTo(IExcelRange range, bool includeFont = true, bool includeBorder = true, bool includeFill = true)
    {
        if (_style == null || range == null)
            return;

        try
        {
            var excelRange = range as ExcelRange;
            if (excelRange?.InternalRange != null)
            {
                // 应用样式到区域
                excelRange.InternalRange.Style = _style.Name;

                // 如果需要应用特定属性，则手动设置
                if (includeFont)
                {
                    excelRange.Font.Name = Font.Name;
                    excelRange.Font.Size = Font.Size;
                    excelRange.Font.Bold = Font.Bold;
                    excelRange.Font.Italic = Font.Italic;
                    excelRange.Font.Color = Font.Color;
                    excelRange.Font.Underline = Font.Underline;
                }



                excelRange.NumberFormat = NumberFormat;
                excelRange.HorizontalAlignment = HorizontalAlignment;
                excelRange.VerticalAlignment = VerticalAlignment;
                excelRange.WrapText = WrapText;
                excelRange.IndentLevel = IndentLevel;
                excelRange.Orientation = Orientation;
                excelRange.Locked = Locked;
                excelRange.FormulaHidden = FormulaHidden;
            }
        }
        catch (Exception ex)
        {
            log.Error($"应用样式 {Name} 到区域 {range.Address} 时发生异常", ex);
        }
    }

    /// <summary>
    /// 重置样式为默认值
    /// </summary>
    public void Reset()
    {
        if (_style == null) return;

        try
        {
            // 重置样式属性为默认值
            NumberFormat = "General";
            HorizontalAlignment = XlHAlign.xlHAlignLeft; // xlLeft
            VerticalAlignment = XlVAlign.xlVAlignJustify;   // xlTop
            WrapText = false;
            IndentLevel = 0;
            Orientation = 0;
            ShrinkToFit = false;
            MergeCells = false;
            Locked = true;
            FormulaHidden = false;

            // 重置字体属性
            Font.Name = "Calibri";
            Font.Size = 11;
            Font.Bold = false;
            Font.Italic = false;
            Font.Underline = 0; // xlUnderlineStyleNone
            Font.Color = Color.Black; // 黑色

            // 重置填充属性
            Interior.Color = Color.White; // 白色
            Interior.Pattern = -4142;  // xlPatternAutomatic
            Interior.PatternColor = Color.Black; // 黑色
        }
        catch (Exception ex)
        {
            log.Error($"重置样式 {Name} 时发生异常", ex);
        }
    }

    #endregion

    #region 高级功能
    /// <summary>
    /// 克隆样式
    /// </summary>
    /// <returns>克隆的样式对象</returns>
    public IExcelStyle? Clone()
    {
        if (_style?.Parent == null)
            return null;

        try
        {
            string cloneName = $"{Name}_Clone_{DateTime.Now:HHmmss}";
            return Copy(cloneName);
        }
        catch (Exception ex)
        {
            log.Error($"克隆样式 {Name} 时发生异常", ex);
            return null;
        }
    }
    #endregion
}
