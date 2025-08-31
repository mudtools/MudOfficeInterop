//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Font 的实现类。
/// </summary>
internal class WordFont : IWordFont
{
    private MsWord.Font _font;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="font">原始 COM Font 对象。</param>
    internal WordFont(MsWord.Font font)
    {
        _font = font ?? throw new ArgumentNullException(nameof(font));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _font != null ? new WordApplication(_font.Application) : null;

    /// <inheritdoc/>
    public object Parent => _font?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _font?.Name ?? string.Empty;
        set
        {
            if (_font != null)
                _font.Name = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _font?.Size ?? 0f;
        set
        {
            if (_font != null)
                _font.Size = value;
        }
    }

    /// <inheritdoc/>
    public bool Bold
    {
        get => _font?.Bold == 1;
        set
        {
            if (_font != null)
                _font.Bold = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool Italic
    {
        get => _font?.Italic == 1;
        set
        {
            if (_font != null)
                _font.Italic = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool Underline
    {
        get => _font?.Underline != MsWord.WdUnderline.wdUnderlineNone;
        set
        {
            if (_font != null)
                _font.Underline = value
                    ? MsWord.WdUnderline.wdUnderlineSingle
                    : MsWord.WdUnderline.wdUnderlineNone;
        }
    }

    /// <inheritdoc/>
    public WdColor Color
    {
        get => (WdColor)(int)_font?.Color;
        set
        {
            if (_font != null)
                _font.Color = (MsWord.WdColor)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool Superscript
    {
        get => _font?.Superscript == 1;
        set
        {
            if (_font != null)
                _font.Superscript = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool Subscript
    {
        get => _font?.Subscript == 1;
        set
        {
            if (_font != null)
                _font.Subscript = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public float Spacing
    {
        get => _font?.Spacing ?? 0f;
        set
        {
            if (_font != null)
                _font.Spacing = value;
        }
    }

    /// <inheritdoc/>
    public int Scaling
    {
        get => _font?.Scaling ?? 100;
        set
        {
            if (_font != null)
                _font.Scaling = value;
        }
    }

    /// <inheritdoc/>
    public int Position
    {
        get => _font?.Position ?? 0;
        set
        {
            if (_font != null)
                _font.Position = value;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _font != null)
        {
            Marshal.ReleaseComObject(_font);
            _font = null;
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