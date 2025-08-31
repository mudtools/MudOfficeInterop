//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ConditionalStyle 的实现类。
/// </summary>
internal class WordConditionalStyle : IWordConditionalStyle
{
    private MsWord.ConditionalStyle _conditionalStyle;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="conditionalStyle">原始 COM ConditionalStyle 对象。</param>
    internal WordConditionalStyle(MsWord.ConditionalStyle conditionalStyle)
    {
        _conditionalStyle = conditionalStyle ?? throw new ArgumentNullException(nameof(conditionalStyle));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _conditionalStyle != null ? new WordApplication(_conditionalStyle.Application) : null;

    /// <inheritdoc/>
    public object Parent => _conditionalStyle?.Parent;

    /// <inheritdoc/>
    public IWordBorders Borders =>
        _conditionalStyle?.Borders != null ? new WordBorders(_conditionalStyle.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading =>
        _conditionalStyle?.Shading != null ? new WordShading(_conditionalStyle.Shading) : null;

    /// <inheritdoc/>
    public IWordFont Font =>
        _conditionalStyle?.Font != null ? new WordFont(_conditionalStyle.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat =>
        _conditionalStyle?.ParagraphFormat != null ? new WordParagraphFormat(_conditionalStyle.ParagraphFormat) : null;

    public float BottomPadding
    {
        get => _conditionalStyle?.BottomPadding ?? 0f;
        set
        {
            if (_conditionalStyle != null)
                _conditionalStyle.BottomPadding = value;
        }
    }

    public float TopPadding
    {
        get => _conditionalStyle?.TopPadding ?? 0f;
        set
        {
            if (_conditionalStyle != null)
                _conditionalStyle.TopPadding = value;
        }
    }

    public float LeftPadding
    {
        get => _conditionalStyle?.LeftPadding ?? 0f;
        set
        {
            if (_conditionalStyle != null)
                _conditionalStyle.LeftPadding = value;
        }
    }

    public float RightPadding
    {
        get => _conditionalStyle?.RightPadding ?? 0f;
        set
        {
            if (_conditionalStyle != null)
                _conditionalStyle.RightPadding = value;
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

        if (disposing)
        {
            // 释放边框集合
            if (_conditionalStyle?.Borders != null)
            {
                Marshal.ReleaseComObject(_conditionalStyle.Borders);
            }
            // 释放底纹对象
            if (_conditionalStyle?.Shading != null)
            {
                Marshal.ReleaseComObject(_conditionalStyle.Shading);
            }
            // 释放字体对象
            if (_conditionalStyle?.Font != null)
            {
                Marshal.ReleaseComObject(_conditionalStyle.Font);
            }
            // 释放段落格式对象
            if (_conditionalStyle?.ParagraphFormat != null)
            {
                Marshal.ReleaseComObject(_conditionalStyle.ParagraphFormat);
            }
            // 释放条件样式对象本身
            if (_conditionalStyle != null)
            {
                Marshal.ReleaseComObject(_conditionalStyle);
                _conditionalStyle = null;
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