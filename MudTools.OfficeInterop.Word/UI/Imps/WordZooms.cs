//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordZooms"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.Zooms 对象。
/// </summary>
/// <remarks>
/// <para>
/// 此类封装了 <c>Microsoft.Office.Interop.Word.Zooms</c> 对象。
/// 请注意，<c>Zooms</c> 集合在 Word 对象模型中没有公开可用的实例，
/// 其功能主要通过关联的 <c>View.Zoom</c> 属性（返回单个 <c>Zoom</c> 对象）来实现。
/// </para>
/// <para>
/// 因此，此类的实现仅限于封装其继承的属性。
/// </para>
/// </remarks>
internal class WordZooms : IWordZooms
{
    private MsWord.Zooms _zooms; // 原始的 COM 对象
    private bool _disposedValue = false; // 用于检测冗余的 Dispose 调用

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordZooms"/> 类的新实例。
    /// </summary>
    /// <param name="zooms">原始的 Microsoft.Office.Interop.Word.Zooms 对象。</param>
    /// <exception cref="ArgumentNullException">如果 <paramref name="zooms"/> 为 null。</exception>
    internal WordZooms(MsWord.Zooms zooms)
    {
        _zooms = zooms ?? throw new ArgumentNullException(nameof(zooms));
    }

    #region 属性实现

    /// <inheritdoc/>
    /// <remarks>此属性继承自 _IMsoDispObj。</remarks>
    public IWordApplication Application => _zooms?.Application != null ? new WordApplication(_zooms.Application) : null;

    /// <inheritdoc/>
    /// <remarks>此属性继承自 _IMsoDispObj。</remarks>
    public int Creator => _zooms?.Creator ?? 0;

    /// <inheritdoc/>
    /// <remarks>对于 Zooms 集合，父对象通常是关联的 View 对象。</remarks>
    public object? Parent => _zooms?.Parent;

    #endregion // 属性实现

    public IWordZoom Item(WdViewType Index)
    {
        var zoom = _zooms.Item((MsWord.WdViewType)(int)Index);
        return zoom != null ? new WordZoom(zoom) : null;
    }

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordZooms"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_zooms != null)
                {
                    Marshal.ReleaseComObject(_zooms);
                    _zooms = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordZooms"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}