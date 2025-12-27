//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordZoom"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.Zoom 对象。
/// </summary>
internal class WordZoom : IWordZoom
{
    private MsWord.Zoom _zoom;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordZoom"/> 类的新实例。
    /// </summary>
    /// <param name="zoom">原始的 Microsoft.Office.Interop.Word.Zoom 对象。</param>
    internal WordZoom(MsWord.Zoom zoom)
    {
        _zoom = zoom ?? throw new ArgumentNullException(nameof(zoom));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _zoom?.Application != null ? new WordApplication(_zoom.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _zoom?.Parent;

    /// <inheritdoc/>
    /// <remarks>
    /// 根据 Word API 文档，设置的值应在 10 到 500 之间。超出此范围可能导致错误或被自动调整。
    /// </remarks>
    public int Percentage
    {
        get => _zoom?.Percentage ?? 100;
        set
        {
            if (_zoom != null)
            {
                _zoom.Percentage = value;
            }
        }
    }

    public int PageRows
    {
        get => _zoom?.PageRows ?? 1;
        set
        {
            if (_zoom != null)
            {
                _zoom.PageRows = value;
            }
        }
    }

    public int PageColumns
    {
        get => _zoom?.PageColumns ?? 1;
        set
        {
            if (_zoom != null)
            {
                _zoom.PageColumns = value;
            }
        }
    }

    /// <inheritdoc/>
    public WdPageFit PageFit
    {
        get => _zoom?.PageFit != null ? (WdPageFit)(int)_zoom?.PageFit : WdPageFit.wdPageFitFullPage;
        set
        {
            if (_zoom != null) _zoom.PageFit = (MsWord.WdPageFit)(int)value;
        }
    }

    #endregion // 属性实现

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordZoom"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_zoom != null)
                {
                    Marshal.ReleaseComObject(_zoom);
                    _zoom = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordZoom"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}