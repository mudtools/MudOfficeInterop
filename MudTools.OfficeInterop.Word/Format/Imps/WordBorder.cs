//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Border 的实现类。
/// </summary>
internal class WordBorder : IWordBorder
{
    private MsWord.Border _border;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="border">原始 COM Border 对象。</param>
    internal WordBorder(MsWord.Border border)
    {
        _border = border ?? throw new ArgumentNullException(nameof(border));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _border != null ? new WordApplication(_border.Application) : null;

    /// <inheritdoc/>
    public object Parent => _border?.Parent;

    /// <inheritdoc/>
    public bool Visible
    {
        get => _border != null && _border.Visible;
        set
        {
            if (_border != null)
                _border.Visible = value;
        }
    }

    /// <inheritdoc/>
    public WdLineStyle LineStyle
    {
        get => (WdLineStyle)(int)_border!.LineStyle;
        set
        {
            if (_border != null)
                _border.LineStyle = (MsWord.WdLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLineWidth LineWidth
    {
        get => (WdLineWidth)(int)_border?.LineWidth;
        set
        {
            if (_border != null)
                _border.LineWidth = (MsWord.WdLineWidth)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColor Color
    {
        get => (WdColor)(int)_border?.Color;
        set
        {
            if (_border != null)
                _border.Color = (MsWord.WdColor)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColorIndex ColorIndex
    {
        get => (WdColorIndex)(int)_border?.ColorIndex;
        set
        {
            if (_border != null)
                _border.ColorIndex = (MsWord.WdColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdPageBorderArt ArtStyle
    {
        get => (WdPageBorderArt)(int)_border?.ArtStyle;
        set
        {
            if (_border != null)
                _border.ArtStyle = (MsWord.WdPageBorderArt)(int)value;
        }
    }

    /// <inheritdoc/>
    public int ArtWidth
    {
        get => _border?.ArtWidth ?? 0;
        set
        {
            if (_border != null)
                _border.ArtWidth = value;
        }
    }

    /// <inheritdoc/>
    public bool Inside => _border != null && _border.Inside;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _border != null)
        {
            Marshal.ReleaseComObject(_border);
            _border = null;
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