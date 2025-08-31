//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ParagraphFormat 的实现类。
/// </summary>
internal class WordParagraphFormat : IWordParagraphFormat
{
    private MsWord.ParagraphFormat _format;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="format">原始 COM ParagraphFormat 对象。</param>
    internal WordParagraphFormat(MsWord.ParagraphFormat format)
    {
        _format = format ?? throw new ArgumentNullException(nameof(format));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _format != null ? new WordApplication(_format.Application) : null;

    /// <inheritdoc/>
    public object Parent => _format?.Parent;

    /// <inheritdoc/>
    public WdParagraphAlignment Alignment
    {
        get => (WdParagraphAlignment)(int)_format?.Alignment;
        set
        {
            if (_format != null)
                _format.Alignment = (MsWord.WdParagraphAlignment)(int)value;
        }
    }

    /// <inheritdoc/>
    public float FirstLineIndent
    {
        get => _format?.FirstLineIndent ?? 0f;
        set
        {
            if (_format != null)
                _format.FirstLineIndent = value;
        }
    }

    /// <inheritdoc/>
    public float LeftIndent
    {
        get => _format?.LeftIndent ?? 0f;
        set
        {
            if (_format != null)
                _format.LeftIndent = value;
        }
    }

    /// <inheritdoc/>
    public float RightIndent
    {
        get => _format?.RightIndent ?? 0f;
        set
        {
            if (_format != null)
                _format.RightIndent = value;
        }
    }

    /// <inheritdoc/>
    public float SpaceBefore
    {
        get => _format?.SpaceBefore ?? 0f;
        set
        {
            if (_format != null)
                _format.SpaceBefore = value;
        }
    }

    /// <inheritdoc/>
    public float SpaceAfter
    {
        get => _format?.SpaceAfter ?? 0f;
        set
        {
            if (_format != null)
                _format.SpaceAfter = value;
        }
    }

    /// <inheritdoc/>
    public WdLineSpacing LineSpacingRule
    {
        get => (WdLineSpacing)(int)_format?.LineSpacingRule;
        set
        {
            if (_format != null)
                _format.LineSpacingRule = (MsWord.WdLineSpacing)(int)value;
        }
    }

    /// <inheritdoc/>
    public float LineSpacing
    {
        get => _format?.LineSpacing ?? 0f;
        set
        {
            if (_format != null)
                _format.LineSpacing = value;
        }
    }

    /// <inheritdoc/>
    public bool WidowControl
    {
        get => _format?.WidowControl == 1;
        set
        {
            if (_format != null)
                _format.WidowControl = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool KeepTogether
    {
        get => _format?.KeepTogether == 1;
        set
        {
            if (_format != null)
                _format.KeepTogether = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool KeepWithNext
    {
        get => _format?.KeepWithNext == 1;
        set
        {
            if (_format != null)
                _format.KeepWithNext = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public IWordTabStops? TabStops =>
        _format?.TabStops != null ? new WordTabStops(_format.TabStops) : null;

    /// <inheritdoc/>
    public WdOutlineLevel OutlineLevel
    {
        get => (WdOutlineLevel)(int)_format?.OutlineLevel;
        set
        {
            if (_format != null)
                _format.OutlineLevel = (MsWord.WdOutlineLevel)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordBorders? Borders => _format?.Borders != null ? new WordBorders(_format.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading => _format?.Shading != null ? new WordShading(_format.Shading) : null;

    /// <inheritdoc/>
    public WdReadingOrder ReadingOrder
    {
        get => (WdReadingOrder)(int)_format?.ReadingOrder;
        set
        {
            if (_format != null)
                _format.ReadingOrder = (MsWord.WdReadingOrder)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool CharacterUnitLeftIndent
    {
        get => _format?.CharacterUnitLeftIndent == 1;
        set
        {
            if (_format != null)
                _format.CharacterUnitLeftIndent = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool CharacterUnitFirstLineIndent
    {
        get => _format?.CharacterUnitFirstLineIndent == 1;
        set
        {
            if (_format != null)
                _format.CharacterUnitFirstLineIndent = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool CharacterUnitRightIndent
    {
        get => _format?.CharacterUnitRightIndent == 1;
        set
        {
            if (_format != null)
                _format.CharacterUnitRightIndent = value ? 1 : 0;
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
            // 释放制表符集合
            if (_format?.TabStops != null)
            {
                Marshal.ReleaseComObject(_format.TabStops);
            }
            // 释放段落格式对象本身
            if (_format != null)
            {
                Marshal.ReleaseComObject(_format);
                _format = null;
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