//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Interior 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Interior 对象的安全访问和资源管理
/// </summary>
internal class ExcelInterior : IExcelInterior
{
    /// <summary>
    /// 底层的 COM Interior 对象
    /// </summary>
    internal MsExcel.Interior _interior;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelInterior 实例
    /// </summary>
    /// <param name="interior">底层的 COM Interior 对象</param>
    internal ExcelInterior(MsExcel.Interior interior)
    {
        _interior = interior ?? throw new ArgumentNullException(nameof(interior));
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
                // 释放底层COM对象
                if (_interior != null)
                    Marshal.ReleaseComObject(_interior);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _interior = null;
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
    /// 获取或设置内部颜色
    /// </summary>
    public Color Color
    {
        get => _interior != null ? ColorTranslator.FromOle(Convert.ToInt32(_interior.Color)) : Color.Black;
        set
        {
            if (_interior != null)
                _interior.Color = ColorTranslator.ToOle(value);
        }
    }

    /// <summary>
    /// 获取或设置内部图案类型
    /// </summary>
    public XlPattern Pattern
    {
        get => _interior != null ? _interior.Pattern.ObjectConvertEnum(XlPattern.xlPatternAutomatic) : XlPattern.xlPatternAutomatic;
        set
        {
            if (_interior != null)
                _interior.Pattern = value.EnumConvert(MsExcel.XlPattern.xlPatternAutomatic);
        }
    }

    public XlColorIndex ColorIndex
    {
        get => _interior != null ? _interior.ColorIndex.ObjectConvertEnum(XlColorIndex.xlColorIndexAutomatic) : XlColorIndex.xlColorIndexAutomatic;
        set
        {
            if (_interior != null)
                _interior.ColorIndex = value.EnumConvert(MsExcel.XlColorIndex.xlColorIndexAutomatic);
        }
    }

    /// <summary>
    /// 获取或设置图案颜色
    /// </summary>
    public Color PatternColor
    {
        get => _interior != null ? ColorTranslator.FromOle(Convert.ToInt32(_interior.PatternColor)) : Color.Black;
        set
        {
            if (_interior != null)
                _interior.PatternColor = ColorTranslator.ToOle(value);
        }
    }

    /// <summary>
    /// 获取或设置主题颜色
    /// </summary>
    public Color ThemeColor
    {
        get => _interior != null ? ColorTranslator.FromOle(Convert.ToInt32(_interior.ThemeColor)) : Color.Black;
        set
        {
            if (_interior != null)
                _interior.ThemeColor = ColorTranslator.ToOle(value);
        }
    }

    /// <summary>
    /// 获取或设置着色和阴影
    /// </summary>
    public double TintAndShade
    {
        get => _interior != null ? Convert.ToDouble(_interior.TintAndShade) : 0;
        set
        {
            if (_interior != null)
                _interior.TintAndShade = value;
        }
    }

    /// <summary>
    /// 获取或设置图案主题颜色
    /// </summary>
    public int PatternThemeColor
    {
        get => _interior != null ? Convert.ToInt32(_interior.PatternThemeColor) : 0;
        set
        {
            if (_interior != null)
                _interior.PatternThemeColor = value;
        }
    }

    /// <summary>
    /// 获取或设置图案着色和阴影
    /// </summary>
    public double PatternTintAndShade
    {
        get => _interior != null ? Convert.ToDouble(_interior.PatternTintAndShade) : 0;
        set
        {
            if (_interior != null)
                _interior.PatternTintAndShade = value;
        }
    }


    /// <summary>
    /// 获取内部对象所在的父对象
    /// </summary>
    public object? Parent => _interior?.Parent;

    /// <summary>
    /// 获取内部对象所在的Application对象
    /// </summary>
    public IExcelApplication? Application
    {
        get
        {
            var application = _interior?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }
    #endregion

    #region 格式设置
    /// <summary>
    /// 重置内部对象为默认值
    /// </summary>
    public void Reset()
    {
        if (_interior == null) return;

        try
        {
            Color = Color.White; // 白色
            Pattern = XlPattern.xlPatternAutomatic;   // xlPatternAutomatic
            PatternColor = Color.Black; // 黑色
            ThemeColor = Color.Black;   // 默认主题颜色
            TintAndShade = 0; // 无着色和阴影
            PatternThemeColor = 0; // 默认图案主题颜色
            PatternTintAndShade = 0; // 无图案着色和阴影
        }
        catch
        {
            // 忽略重置过程中的异常
        }
    }

    /// <summary>
    /// 复制内部对象格式
    /// </summary>
    /// <param name="sourceInterior">源内部对象</param>
    public void CopyFormat(IExcelInterior sourceInterior)
    {
        if (_interior == null || sourceInterior == null) return;

        try
        {
            Color = sourceInterior.Color;
            Pattern = sourceInterior.Pattern;
            PatternColor = sourceInterior.PatternColor;
            ThemeColor = sourceInterior.ThemeColor;
            TintAndShade = sourceInterior.TintAndShade;
            PatternThemeColor = sourceInterior.PatternThemeColor;
            PatternTintAndShade = sourceInterior.PatternTintAndShade;
        }
        catch
        {
            // 忽略复制格式过程中的异常
        }
    }
    #endregion
}