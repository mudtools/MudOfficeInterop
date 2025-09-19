//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordView"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.View 对象。
/// </summary>
internal class WordView : IWordView
{
    private MsWord.View _view;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordView"/> 类的新实例。
    /// </summary>
    /// <param name="view">原始的 Microsoft.Office.Interop.Word.View 对象。</param>
    internal WordView(MsWord.View view)
    {
        _view = view ?? throw new ArgumentNullException(nameof(view));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _view?.Application != null ? new WordApplication(_view.Application) : null;

    /// <inheritdoc/>
    public object Parent => _view?.Parent;

    /// <inheritdoc/>
    public WdViewType Type
    {
        get => _view?.Type != null ? _view.Type.EnumConvert(WdViewType.wdNormalView) : WdViewType.wdNormalView;
        set
        {
            if (_view != null) _view.Type = value.EnumConvert(MsWord.WdViewType.wdNormalView);
        }
    }

    /// <inheritdoc/>
    public WdSeekView SeekView
    {
        get => _view?.SeekView != null ? _view.SeekView.EnumConvert(WdSeekView.wdSeekMainDocument) : WdSeekView.wdSeekMainDocument;
        set
        {
            if (_view != null) _view.SeekView = value.EnumConvert(MsWord.WdSeekView.wdSeekMainDocument);
        }
    }

    /// <inheritdoc/>
    public bool ShowParagraphs
    {
        get => _view?.ShowParagraphs ?? false;
        set { if (_view != null) _view.ShowParagraphs = value; }
    }

    /// <inheritdoc/>
    public bool ShowHyphens
    {
        get => _view?.ShowHyphens ?? false;
        set { if (_view != null) _view.ShowHyphens = value; }
    }

    /// <inheritdoc/>
    public bool ShowHiddenText
    {
        get => _view?.ShowHiddenText ?? false;
        set { if (_view != null) _view.ShowHiddenText = value; }
    }

    /// <inheritdoc/>
    public bool ShowBookmarks
    {
        get => _view?.ShowBookmarks ?? false;
        set { if (_view != null) _view.ShowBookmarks = value; }
    }

    /// <inheritdoc/>
    public bool ShowObjectAnchors
    {
        get => _view?.ShowObjectAnchors ?? false;
        set { if (_view != null) _view.ShowObjectAnchors = value; }
    }

    /// <inheritdoc/>
    public bool ShowTextBoundaries
    {
        get => _view?.ShowTextBoundaries ?? false;
        set { if (_view != null) _view.ShowTextBoundaries = value; }
    }

    /// <inheritdoc/>
    public bool ShowHighlight
    {
        get => _view?.ShowHighlight ?? true;
        set { if (_view != null) _view.ShowHighlight = value; }
    }

    /// <inheritdoc/>
    public bool ShowFieldCodes
    {
        get => _view?.ShowFieldCodes ?? false;
        set { if (_view != null) _view.ShowFieldCodes = value; }
    }

    /// <inheritdoc/>
    public bool ShowTabs
    {
        get => _view?.ShowTabs ?? false;
        set { if (_view != null) _view.ShowTabs = value; }
    }

    /// <inheritdoc/>
    public bool ShowSpaces
    {
        get => _view?.ShowSpaces ?? false;
        set { if (_view != null) _view.ShowSpaces = value; }
    }

    /// <inheritdoc/>
    public bool ShowAll
    {
        get => _view?.ShowAll ?? false;
        set { if (_view != null) _view.ShowAll = value; }
    }

    /// <inheritdoc/>
    public bool ShowMainTextLayer
    {
        get => _view?.ShowMainTextLayer ?? true;
        set { if (_view != null) _view.ShowMainTextLayer = value; }
    }

    /// <inheritdoc/>
    public bool ShowComments
    {
        get => _view?.ShowComments ?? true;
        set { if (_view != null) _view.ShowComments = value; }
    }

    /// <inheritdoc/>
    public bool ShowInkAnnotations
    {
        get => _view?.ShowInkAnnotations ?? true;
        set { if (_view != null) _view.ShowInkAnnotations = value; }
    }

    /// <inheritdoc/>
    public int ShowXMLMarkup
    {
        get => _view?.ShowXMLMarkup ?? 0;
        set { if (_view != null) _view.ShowXMLMarkup = value; }
    }

    /// <inheritdoc/>
    public bool ShowFormatChanges
    {
        get => _view?.ShowFormatChanges ?? true;
        set { if (_view != null) _view.ShowFormatChanges = value; }
    }

    /// <inheritdoc/>
    public bool ShowRevisionsAndComments
    {
        get => _view?.ShowRevisionsAndComments ?? true;
        set { if (_view != null) _view.ShowRevisionsAndComments = value; }
    }

    /// <inheritdoc/>
    public WdRevisionsView RevisionsView
    {
        get => _view?.RevisionsView != null ? _view.RevisionsView.EnumConvert(WdRevisionsView.wdRevisionsViewFinal) : WdRevisionsView.wdRevisionsViewFinal;
        set
        {
            if (_view != null) _view.RevisionsView = value.EnumConvert(MsWord.WdRevisionsView.wdRevisionsViewFinal);
        }
    }

    /// <inheritdoc/>
    public WdRevisionsMode RevisionsMode
    {
        get => _view?.RevisionsMode != null ? _view.RevisionsMode.EnumConvert(WdRevisionsMode.wdBalloonRevisions) : WdRevisionsMode.wdBalloonRevisions;
        set
        {
            if (_view != null) _view.RevisionsMode = value.EnumConvert(MsWord.WdRevisionsMode.wdBalloonRevisions);
        }
    }

    /// <inheritdoc/>
    public IWordZoom Zoom => _view?.Zoom != null ? new WordZoom(_view.Zoom) : null;

    /// <inheritdoc/>
    public bool FullScreen
    {
        get => _view?.FullScreen ?? false;
        set { if (_view != null) _view.FullScreen = value; }
    }

    /// <inheritdoc/>
    public bool Magnifier
    {
        get => _view?.Magnifier ?? false;
        set { if (_view != null) _view.Magnifier = value; }
    }

    /// <inheritdoc/>
    public bool Panning
    {
        get => _view?.Panning ?? false;
        set { if (_view != null) _view.Panning = value; }
    }

    /// <inheritdoc/>
    public bool ReadingLayout
    {
        get => _view?.ReadingLayout ?? false;
        set { if (_view != null) _view.ReadingLayout = value; }
    }

    /// <inheritdoc/>
    public bool WrapToWindow
    {
        get => _view?.WrapToWindow ?? false;
        set { if (_view != null) _view.WrapToWindow = value; }
    }

    /// <inheritdoc/>
    public WdRevisionsBalloonWidthType RevisionsBalloonWidthType
    {
        get => _view?.RevisionsBalloonWidthType != null ? _view.RevisionsBalloonWidthType.EnumConvert(WdRevisionsBalloonWidthType.wdBalloonWidthPercent) : WdRevisionsBalloonWidthType.wdBalloonWidthPercent;
        set
        {
            if (_view != null) _view.RevisionsBalloonWidthType = value.EnumConvert(MsWord.WdRevisionsBalloonWidthType.wdBalloonWidthPercent);
        }
    }

    /// <inheritdoc/>
    public float RevisionsBalloonWidth
    {
        get => _view?.RevisionsBalloonWidth ?? 0.0f;
        set { if (_view != null) _view.RevisionsBalloonWidth = value; }
    }


    #endregion // 属性实现

    #region 方法实现

    /// <inheritdoc/>
    public void CollapseAllHeadings()
    {
        _view?.CollapseAllHeadings();
    }

    /// <inheritdoc/>
    public void ExpandAllHeadings()
    {
        _view?.ExpandAllHeadings();
    }

    /// <inheritdoc/>
    public void ExpandOutline(IWordRange rang)
    {
        _view?.ExpandOutline(((WordRange)rang)._range);
    }

    /// <inheritdoc/>
    public void CollapseOutline(IWordRange rang)
    {
        _view?.CollapseOutline(((WordRange)rang)._range);
    }

    /// <inheritdoc/>
    public void NextHeaderFooter()
    {
        _view?.NextHeaderFooter();
    }

    /// <inheritdoc/>
    public void PreviousHeaderFooter()
    {
        _view?.PreviousHeaderFooter();
    }

    /// <inheritdoc/>
    public void ShowAllHeadings()
    {
        _view?.ShowAllHeadings();
    }

    /// <inheritdoc/>
    public void ShowHeading(int headingLevel)
    {
        if (headingLevel < 1 || headingLevel > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(headingLevel), "Heading level must be between 1 and 9.");
        }
        _view?.ShowHeading(headingLevel);
    }

    #endregion // 方法实现

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordView"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_view != null)
                {
                    Marshal.ReleaseComObject(_view);
                    _view = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordView"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
#endregion