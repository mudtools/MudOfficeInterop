//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.SmartArtLayouts 的二次封装实现类。
/// 提供安全访问 SmartArt 布局集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeSmartArtLayouts : IOfficeSmartArtLayouts
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeSmartArtLayouts));
    private MsCore.SmartArtLayouts _smartArtLayouts;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 SmartArtLayouts 对象。
    /// </summary>
    /// <param name="smartArtLayouts">原始的 COM SmartArtLayouts 对象。</param>
    internal OfficeSmartArtLayouts(MsCore.SmartArtLayouts smartArtLayouts)
    {
        _smartArtLayouts = smartArtLayouts ?? throw new ArgumentNullException(nameof(smartArtLayouts));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _smartArtLayouts?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeSmartArtLayout this[int index]
    {
        get
        {
            if (_smartArtLayouts == null || index < 1 || index > Count)
                return null;

            try
            {
                var layout = _smartArtLayouts[index];
                return layout != null ? new OfficeSmartArtLayout(layout) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IOfficeSmartArtLayout this[string id]
    {
        get
        {
            if (_smartArtLayouts == null || string.IsNullOrEmpty(id))
                return null;

            try
            {
                var layout = _smartArtLayouts[id];
                return layout != null ? new OfficeSmartArtLayout(layout) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public object Parent => _smartArtLayouts?.Parent;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficeSmartArtLayout FindByName(string layoutName)
    {
        if (_smartArtLayouts == null || string.IsNullOrEmpty(layoutName))
            return null;

        try
        {
            // 遍历查找匹配名称的布局
            for (int i = 1; i <= Count; i++)
            {
                var layout = _smartArtLayouts[i];
                if (layout != null && string.Equals(layout.Name, layoutName, StringComparison.OrdinalIgnoreCase))
                {
                    return new OfficeSmartArtLayout(layout);
                }
            }
            return null;
        }
        catch (COMException ce)
        {
            log.Error($"Failed to find object by name: {ce.Message}", ce);
            return null;
        }
    }

    #endregion

    #region IEnumerable<IOfficeSmartArtLayout> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeSmartArtLayout> GetEnumerator()
    {
        if (_smartArtLayouts == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var layout = _smartArtLayouts[i];
            if (layout != null)
                yield return new OfficeSmartArtLayout(layout);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _smartArtLayouts != null)
        {
            try
            {
                Marshal.ReleaseComObject(_smartArtLayouts);
            }
            catch
            {
                // 忽略释放异常
            }
            _smartArtLayouts = null;
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