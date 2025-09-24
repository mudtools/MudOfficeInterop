
using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// FormatColor（实际为 ColorFormat）COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelFormatColor : IExcelFormatColor
{
    internal MsExcel.FormatColor _colorFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="colorFormat">原始的 FormatColor COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 colorFormat 为 null 时抛出。</exception>
    internal ExcelFormatColor(MsExcel.FormatColor colorFormat)
    {
        _colorFormat = colorFormat ?? throw new ArgumentNullException(nameof(colorFormat));
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
            if (_colorFormat != null)
            {
                Marshal.ReleaseComObject(_colorFormat);
                _colorFormat = null;
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
    /// 获取此对象的父对象（如 FillFormat、ShadowFormat 等）。
    /// </summary>
    public object? Parent => _colorFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _colorFormat?.Application != null
            ? new ExcelApplication(_colorFormat.Application)
            : null;

    /// <summary>
    /// 获取或设置颜色的 RGB 值（如 0xFF0000 表示红色）。
    /// 设置此属性会清除主题色设置。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public int RGB
    {
        get => _colorFormat != null ? _colorFormat.Color.ConvertToInt() : 0;

        set
        {
            if (_colorFormat != null)
            {
                _colorFormat.Color = value;
            }
        }
    }

    public Color Color
    {
        get
        {
            if (_colorFormat != null)
            {
                return Color.FromArgb(_colorFormat.Color.ConvertToInt());
            }
            return Color.Empty;
        }
        set
        {
            if (_colorFormat != null)
            {
                _colorFormat.Color = value.ToArgb();
            }
        }
    }

    public XlColorIndex ColorIndex
    {
        get => _colorFormat != null
            ? _colorFormat.ColorIndex.EnumConvert(XlColorIndex.xlColorIndexAutomatic)
            : XlColorIndex.xlColorIndexAutomatic;

        set
        {
            if (_colorFormat != null)
            {
                _colorFormat.ColorIndex = value.EnumConvert(MsExcel.XlColorIndex.xlColorIndexAutomatic);
            }
        }
    }

    /// <summary>
    /// 获取或设置颜色的主题色类型（如强调色1、背景色等）。
    /// 设置此属性会清除 RGB 设置。
    /// 默认值：msoThemeColorMixed
    /// </summary>
    public MsoThemeColorIndex ThemeColor
    {
        get => _colorFormat != null
            ? _colorFormat.ThemeColor.ObjectConvertEnum(MsoThemeColorIndex.msoThemeColorMixed)
            : MsoThemeColorIndex.msoThemeColorMixed;

        set
        {
            if (_colorFormat != null)
            {
                _colorFormat.ThemeColor = value.EnumConvert(MsCore.MsoThemeColorIndex.msoThemeColorMixed);
            }
        }
    }

    /// <summary>
    /// 获取或设置基于主题色的色调调整（-1.0 ~ 1.0）。
    /// 负值变暗，正值变亮，0 表示不调整。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public float TintAndShade
    {
        get => _colorFormat != null ? _colorFormat.TintAndShade.ConvertToFloat() : 0f;

        set
        {
            if (_colorFormat != null)
            {
                if (value < -1.0f || value > 1.0f)
                    throw new ArgumentOutOfRangeException(nameof(value), "TintAndShade 必须在 -1.0 到 1.0 之间。");
                _colorFormat.TintAndShade = value;
            }
        }
    }
}