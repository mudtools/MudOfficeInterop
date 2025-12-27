//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示文档中内容控件的封装实现类。
/// </summary>
internal class WordContentControl : IWordContentControl
{
    private MsWord.ContentControl _contentControl;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordContentControl"/> 类的新实例。
    /// </summary>
    /// <param name="contentControl">要封装的原始 COM ContentControl 对象。</param>
    internal WordContentControl(MsWord.ContentControl contentControl)
    {
        _contentControl = contentControl ?? throw new ArgumentNullException(nameof(contentControl));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _contentControl != null ? new WordApplication(_contentControl.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _contentControl?.Parent;

    /// <inheritdoc/>
    public string Title
    {
        get => _contentControl?.Title ?? string.Empty;
        set { if (_contentControl != null) _contentControl.Title = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string Tag
    {
        get => _contentControl?.Tag ?? string.Empty;
        set { if (_contentControl != null) _contentControl.Tag = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public WdContentControlType Type => _contentControl?.Type != null ? (WdContentControlType)(int)_contentControl?.Type : WdContentControlType.wdContentControlText;

    /// <inheritdoc/>
    public IWordRange Range => _contentControl?.Range != null ? new WordRange(_contentControl.Range) : null;

    /// <inheritdoc/>
    public IWordBuildingBlock PlaceholderText => _contentControl?.PlaceholderText != null ? new WordBuildingBlock(_contentControl.PlaceholderText) : null;

    /// <inheritdoc/>
    public bool LockContentControl
    {
        get => _contentControl?.LockContentControl ?? false;
        set { if (_contentControl != null) _contentControl.LockContentControl = value; }
    }

    /// <inheritdoc/>
    public bool LockContents
    {
        get => _contentControl?.LockContents ?? false;
        set { if (_contentControl != null) _contentControl.LockContents = value; }
    }

    /// <inheritdoc/>
    public bool Temporary
    {
        get => _contentControl?.Temporary ?? false;
        set { if (_contentControl != null) _contentControl.Temporary = value; }
    }

    /// <inheritdoc/>
    public string Text
    {
        get => _contentControl?.Range?.Text ?? string.Empty;
        set { if (_contentControl?.Range != null) _contentControl.Range.Text = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    //public IWordXMLMapping XMLMapping => _contentControl?.XMLMapping;

    /// <inheritdoc/>
    public IWordContentControlListEntries? DropdownListEntries =>
        _contentControl?.DropdownListEntries != null ? new WordContentControlListEntries(_contentControl.DropdownListEntries) : null;

    /// <inheritdoc/>
    public string DateDisplayFormat
    {
        get => _contentControl?.DateDisplayFormat ?? string.Empty;
        set { if (_contentControl != null) _contentControl.DateDisplayFormat = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public WdContentControlDateStorageFormat DateStorageFormat
    {
        get => _contentControl?.DateStorageFormat != null ? (WdContentControlDateStorageFormat)(int)_contentControl?.DateStorageFormat : WdContentControlDateStorageFormat.wdContentControlDateStorageText;
        set
        {
            if (_contentControl != null) _contentControl.DateStorageFormat = (MsWord.WdContentControlDateStorageFormat)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLanguageID DateDisplayLocale
    {
        get => _contentControl?.DateDisplayLocale != null ? (WdLanguageID)(int)_contentControl?.DateDisplayLocale : WdLanguageID.wdSimplifiedChinese;
        set
        {
            if (_contentControl != null) _contentControl.DateDisplayLocale = (MsWord.WdLanguageID)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool ShowingPlaceholderText
    {
        get => _contentControl?.ShowingPlaceholderText ?? false;
    }

    /// <inheritdoc/>
    public bool MultiLine
    {
        get => _contentControl?.MultiLine ?? false;
        set { if (_contentControl != null) _contentControl.MultiLine = value; }
    }

    /// <inheritdoc/>
    public string ID => _contentControl?.ID ?? string.Empty;

    /// <inheritdoc/>
    public bool Checked
    {
        get => _contentControl?.Checked ?? false;
        set { if (_contentControl != null) _contentControl.Checked = value; }
    }

    /// <inheritdoc/>
    public IWordContentControl ParentContentControl => _contentControl?.ParentContentControl != null ? new WordContentControl(_contentControl.ParentContentControl) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete(bool deleteContents)
    {
        _contentControl?.Delete(deleteContents);
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _contentControl?.Copy();
    }

    /// <inheritdoc/>
    public void SetUncheckedSymbol(int characterNumber, string font = "")
    {
        _contentControl?.SetUncheckedSymbol(characterNumber, font);
    }

    /// <inheritdoc/>
    public void SetCheckedSymbol(int characterNumber, string font = "")
    {
        _contentControl?.SetCheckedSymbol(characterNumber, font);
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordContentControl"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _contentControl != null)
        {
            Marshal.ReleaseComObject(_contentControl);
            _contentControl = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordContentControl"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}