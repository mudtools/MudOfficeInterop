//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Microsoft Word 中应用程序设置的封装实现类。
/// </summary>
internal partial class WordOptions : IWordOptions
{
    private MsWord.Options _options;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 WordOptions 类的新实例
    /// </summary>
    /// <param name="options">原始的 Microsoft.Office.Interop.Word.Options 对象</param>
    /// <exception cref="ArgumentNullException">当 options 为 null 时抛出</exception>
    internal WordOptions(MsWord.Options options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _disposedValue = false;
    }

    #region 属性实现

    public bool AllowDragAndDrop
    {
        get
        {
            return _options.AllowDragAndDrop;
        }
        set
        {
            _options.AllowDragAndDrop = value;
        }
    }

    public bool AutoCreateNewDrawings
    {
        get
        {
            return _options.AutoCreateNewDrawings;
        }
        set
        {
            _options.AutoCreateNewDrawings = value;
        }
    }

    public bool EnableLivePreview
    {
        get
        {
            return _options.EnableLivePreview;
        }
        set
        {
            _options.EnableLivePreview = value;
        }
    }


    #endregion

    #region 视图与显示属性实现

    public bool DisplayGridLines
    {
        get
        {
            return _options.DisplayGridLines;
        }
        set
        {
            _options.DisplayGridLines = value;
        }
    }
    #endregion

    #region 编辑与输入选项实现

    public bool ReplaceSelection
    {
        get
        {
            return _options.ReplaceSelection;
        }
        set
        {
            _options.ReplaceSelection = value;
        }
    }

    public bool AutoFormatAsYouTypeApplyHeadings
    {
        get
        {
            return _options.AutoFormatAsYouTypeApplyHeadings;
        }
        set
        {
            _options.AutoFormatAsYouTypeApplyHeadings = value;
        }
    }

    #endregion

    #region 保存与备份选项实现

    public bool SaveNormalPrompt
    {
        get
        {
            return _options.SaveNormalPrompt;
        }
        set
        {
            _options.SaveNormalPrompt = value;
        }
    }

    public bool SavePropertiesPrompt
    {
        get
        {
            return _options.SavePropertiesPrompt;
        }
        set
        {
            _options.SavePropertiesPrompt = value;
        }
    }

    #endregion

    #region 打印与输出选项实现

    public bool UpdateFieldsAtPrint
    {
        get
        {
            return _options.UpdateFieldsAtPrint;
        }
        set
        {
            _options.UpdateFieldsAtPrint = value;
        }
    }

    public bool UpdateLinksAtPrint
    {
        get
        {
            return _options.UpdateLinksAtPrint;
        }
        set
        {
            _options.UpdateLinksAtPrint = value;
        }
    }

    public bool PrintHiddenText
    {
        get
        {
            return _options.PrintHiddenText;
        }
        set
        {
            _options.PrintHiddenText = value;
        }
    }

    public bool PrintDraft
    {
        get
        {
            return _options.PrintDraft;
        }
        set
        {
            _options.PrintDraft = value;
        }
    }

    public bool PrintReverse
    {
        get
        {
            return _options.PrintReverse;
        }
        set
        {
            _options.PrintReverse = value;
        }
    }

    #endregion


    #region 语言与校对选项实现

    public bool CheckSpellingAsYouType
    {
        get
        {
            return _options.CheckSpellingAsYouType;
        }
        set
        {
            _options.CheckSpellingAsYouType = value;
        }
    }

    public bool CheckGrammarAsYouType
    {
        get
        {
            return _options.CheckGrammarAsYouType;
        }
        set
        {
            _options.CheckGrammarAsYouType = value;
        }
    }

    public bool IgnoreUppercase
    {
        get
        {
            return _options.IgnoreUppercase;
        }
        set
        {
            _options.IgnoreUppercase = value;
        }
    }

    public bool IgnoreMixedDigits
    {
        get
        {
            return _options.IgnoreMixedDigits;
        }
        set
        {
            _options.IgnoreMixedDigits = value;
        }
    }

    #endregion

    #region 高级选项实现
    public bool EnableSound
    {
        get
        {
            return _options.EnableSound;
        }
        set
        {
            _options.EnableSound = value;
        }
    }

    #endregion



    #region 修订与跟踪选项实现

    public WdColorIndex InsertedTextColor
    {
        get
        {
            return (WdColorIndex)(int)_options.InsertedTextColor;
        }
        set
        {
            _options.InsertedTextColor = (MsWord.WdColorIndex)(int)value;
        }
    }

    public WdColorIndex DeletedTextColor
    {
        get
        {
            return (WdColorIndex)(int)_options.DeletedTextColor;
        }
        set
        {
            _options.DeletedTextColor = (MsWord.WdColorIndex)(int)value;
        }
    }

    public WdColorIndex RevisedLinesColor
    {
        get
        {
            return (WdColorIndex)(int)_options.RevisedLinesColor;
        }
        set
        {
            _options.RevisedLinesColor = (MsWord.WdColorIndex)(int)value;
        }
    }

    #endregion

    #region 高级编辑选项实现

    public bool CtrlClickHyperlinkToOpen
    {
        get
        {
            return _options.CtrlClickHyperlinkToOpen;
        }
        set
        {
            _options.CtrlClickHyperlinkToOpen = value;
        }
    }

    public bool AllowAccentedUppercase
    {
        get
        {
            return _options.AllowAccentedUppercase;
        }
        set
        {
            _options.AllowAccentedUppercase = value;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordOptions"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _options != null)
        {
            Marshal.ReleaseComObject(_options);
            _options = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordOptions"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}