//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ColorFormat 的实现类。
/// </summary>
internal class WordColorFormat : IWordColorFormat
{
    private MsWord.ColorFormat _colorFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="colorFormat">原始 COM ColorFormat 对象。</param>
    internal WordColorFormat(MsWord.ColorFormat colorFormat)
    {
        _colorFormat = colorFormat ?? throw new ArgumentNullException(nameof(colorFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _colorFormat != null ? new WordApplication(_colorFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _colorFormat?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _colorFormat?.Name ?? string.Empty;
        set
        {
            if (_colorFormat != null)
                _colorFormat.Name = value;
        }
    }

    /// <inheritdoc/>
    public bool OverPrint
    {
        get => _colorFormat?.OverPrint == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_colorFormat != null)
                _colorFormat.OverPrint = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public int RGB
    {
        get => _colorFormat?.RGB ?? 0;
        set
        {
            if (_colorFormat != null)
                _colorFormat.RGB = value;
        }
    }

    /// <inheritdoc/>
    public float Brightness
    {
        get => _colorFormat?.Brightness ?? 0f;
        set
        {
            if (_colorFormat != null)
                _colorFormat.Brightness = value;
        }
    }

    /// <inheritdoc/>
    public int SchemeColor
    {
        get => _colorFormat?.SchemeColor ?? 0;
        set
        {
            if (_colorFormat?.SchemeColor != null)
                _colorFormat.SchemeColor = value;
        }
    }

    /// <inheritdoc/>
    public WdThemeColorIndex ObjectThemeColor
    {
        get => _colorFormat?.ObjectThemeColor != null ? (WdThemeColorIndex)(int)_colorFormat.ObjectThemeColor : WdThemeColorIndex.wdNotThemeColor;
        set
        {
            if (_colorFormat?.ObjectThemeColor != null)
                _colorFormat.ObjectThemeColor = (MsWord.WdThemeColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoColorType Type => _colorFormat?.Type != null ? (MsoColorType)(int)_colorFormat.Type : MsoColorType.msoColorTypeRGB;


    /// <inheritdoc/>
    public float TintAndShade
    {
        get => _colorFormat?.TintAndShade ?? 0f;
        set
        {
            if (_colorFormat != null)
                _colorFormat.TintAndShade = value;
        }
    }
    #endregion

    public void SetCMYK(int Cyan, int Magenta, int Yellow, int Black)
    {
        _colorFormat?.SetCMYK(Cyan, Magenta, Yellow, Black);
    }

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放颜色格式对象本身
            if (_colorFormat != null)
            {
                Marshal.ReleaseComObject(_colorFormat);
                _colorFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}