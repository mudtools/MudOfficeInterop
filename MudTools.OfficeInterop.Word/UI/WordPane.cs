//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordPane"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.Pane 对象。
/// </summary>
internal class WordPane : IWordPane
{
    private MsWord.Pane _pane;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordPane"/> 类的新实例。
    /// </summary>
    /// <param name="pane">原始的 Microsoft.Office.Interop.Word.Pane 对象。</param>
    internal WordPane(MsWord.Pane pane)
    {
        _pane = pane ?? throw new ArgumentNullException(nameof(pane));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _pane?.Application != null ? new WordApplication(_pane.Application) : null;

    /// <inheritdoc/>
    public int? VerticalPercentScrolled
    {
        get => _pane?.VerticalPercentScrolled;
        set { if (_pane != null) _pane.VerticalPercentScrolled = value ?? 0; }
    }

    /// <inheritdoc/>
    public int? Index => _pane?.Index;

    /// <inheritdoc/>
    public int? HorizontalPercentScrolled
    {
        get => _pane?.HorizontalPercentScrolled;
        set { if (_pane != null) _pane.HorizontalPercentScrolled = value ?? 0; }
    }

    /// <inheritdoc/>
    public IWordDocument? Document => _pane?.Document != null ? new WordDocument(_pane.Document) : null;

    /// <inheritdoc/>
    public IWordView? View => _pane?.View != null ? new WordView(_pane.View) : null;

    /// <inheritdoc/>
    public object? Parent => _pane?.Parent;

    /// <inheritdoc/>
    public IWordSelection? Selection => _pane?.Selection != null ? new WordSelection(_pane.Selection) : null;

    /// <inheritdoc/>
    public IWordPages? Pages => _pane?.Pages != null ? new WordPages(_pane.Pages) : null;

    public IWordZooms? Zooms => _pane?.Zooms != null ? new WordZooms(_pane.Zooms) : null;

    #endregion // 属性实现

    #region 方法实现

    /// <inheritdoc/>
    public void Activate()
    {
        _pane?.Activate();
    }

    /// <inheritdoc/>
    public void AutoScroll(int velocity)
    {
        _pane?.AutoScroll(velocity);
    }

    /// <inheritdoc/>
    public void Close()
    {
        _pane?.Close();
    }

    /// <inheritdoc/>
    public void NewFrameset()
    {
        _pane?.NewFrameset();
    }

    /// <inheritdoc/>
    public void LargeScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null)
    {
        object downObj = down ?? (object)missing;
        object upObj = up ?? (object)missing;
        object toRightObj = toRight ?? (object)missing;
        object toLeftObj = toLeft ?? (object)missing;

        _pane?.LargeScroll(
            ref downObj,
            ref upObj,
            ref toRightObj,
            ref toLeftObj
        );
    }

    /// <inheritdoc/>
    public void PageScroll(int? pages = null, int? lines = null)
    {
        object pagesObj = pages ?? (object)missing;
        object linesObj = lines ?? (object)missing;
        _pane?.PageScroll(ref pagesObj, ref linesObj);
    }

    /// <inheritdoc/>
    public void SmallScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null)
    {
        object downObj = down ?? (object)missing;
        object upObj = up ?? (object)missing;
        object toRightObj = toRight ?? (object)missing;
        object toLeftObj = toLeft ?? (object)missing;

        _pane?.SmallScroll(
            ref downObj,
            ref upObj,
            ref toRightObj,
            ref toLeftObj
        );
    }

    #endregion // 方法实现

    #region IDisposable 实现

    // 用于表示缺失的可选参数
    private static readonly object missing = Type.Missing;

    /// <summary>
    /// 释放由 <see cref="WordPane"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_pane != null)
                {
                    Marshal.ReleaseComObject(_pane);
                    _pane = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordPane"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}