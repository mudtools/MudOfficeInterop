//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Style 的实现类。
/// </summary>
internal class WordStyle : IWordStyle
{
    private MsWord.Style _style;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="style">原始 COM Style 对象。</param>
    internal WordStyle(MsWord.Style style)
    {
        _style = style ?? throw new ArgumentNullException(nameof(style));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _style != null ? new WordApplication(_style.Application) : null;

    /// <inheritdoc/>
    public object Parent => _style?.Parent;

    /// <inheritdoc/>
    public bool InUse => _style.InUse;

    /// <inheritdoc/>
    public string NameLocal => _style?.NameLocal ?? string.Empty;

    /// <inheritdoc/>
    public WdStyleType Type
    {
        get => (WdStyleType)(int)_style?.Type;
    }

    /// <inheritdoc/>
    public string NextParagraphStyle
    {
        get => _style?.get_NextParagraphStyle().ToString();
        set
        {
            if (_style != null)
                _style.set_NextParagraphStyle(value);
        }
    }

    /// <inheritdoc/>
    public bool AutomaticallyUpdate
    {
        get => _style.AutomaticallyUpdate;
        set
        {
            if (_style != null)
                _style.AutomaticallyUpdate = value;
        }
    }


    /// <inheritdoc/>
    public bool QuickStyle
    {
        get => _style.QuickStyle;
        set
        {
            if (_style != null)
                _style.QuickStyle = value;
        }
    }

    /// <inheritdoc/>
    public bool Visibility
    {
        get => _style.Visibility;
        set
        {
            if (_style != null)
                _style.Visibility = value;
        }
    }

    /// <inheritdoc/>
    public IWordFont Font => _style?.Font != null ? new WordFont(_style.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat =>
        _style?.ParagraphFormat != null ? new WordParagraphFormat(_style.ParagraphFormat) : null;


    /// <inheritdoc/>
    public IWordListTemplate ListTemplate => _style?.ListTemplate != null ? new WordListTemplate(_style.ListTemplate) : null;

    /// <inheritdoc/>
    public bool IsBuiltIn => _style.BuiltIn;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _style?.Delete();
    }

    /// <inheritdoc/>
    public IWordStyle Copy(string newName)
    {
        if (_style == null)
        {
            throw new ObjectDisposedException(nameof(WordStyle));
        }

        if (string.IsNullOrWhiteSpace(newName))
            throw new ArgumentException("样式名称不能为空。", nameof(newName));

        try
        {
            var newStyle = _style.Application.ActiveDocument.Styles.Add(newName, _style.Type);

            // 复制字体和段落格式
            if (_style.Font != null && newStyle.Font != null)
            {
                newStyle.Font.Name = _style.Font.Name;
                newStyle.Font.Size = _style.Font.Size;
                newStyle.Font.Bold = _style.Font.Bold;
                newStyle.Font.Italic = _style.Font.Italic;
            }

            if (_style.ParagraphFormat != null && newStyle.ParagraphFormat != null)
            {
                newStyle.ParagraphFormat.Alignment = _style.ParagraphFormat.Alignment;
                newStyle.ParagraphFormat.FirstLineIndent = _style.ParagraphFormat.FirstLineIndent;
                newStyle.ParagraphFormat.LeftIndent = _style.ParagraphFormat.LeftIndent;
            }

            return new WordStyle(newStyle);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法复制样式 '{NameLocal}' 到 '{newName}'。", ex);
        }
    }

    /// <inheritdoc/>
    public void ApplyTo(IWordRange range)
    {
        if (_style == null || range == null)
            return;

        try
        {
            var wordRange = (range as WordRange)?._range;
            wordRange?.set_Style(NameLocal);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法将样式 '{NameLocal}' 应用到范围。", ex);
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
            // 释放字体对象
            if (_style?.Font != null)
            {
                Marshal.ReleaseComObject(_style.Font);
            }
            // 释放段落格式对象
            if (_style?.ParagraphFormat != null)
            {
                Marshal.ReleaseComObject(_style.ParagraphFormat);
            }
            // 释放样式对象本身
            if (_style != null)
            {
                Marshal.ReleaseComObject(_style);
                _style = null;
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