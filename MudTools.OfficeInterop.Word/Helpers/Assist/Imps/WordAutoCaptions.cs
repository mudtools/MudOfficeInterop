//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示所有自动图文集条目集合的封装实现类。
/// </summary>
internal class WordAutoCaptions : IWordAutoCaptions
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordAutoCaptions));
    private MsWord.AutoCaptions _autoCaptions;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordAutoCaptions"/> 类的新实例。
    /// </summary>
    /// <param name="autoCaptions">要封装的原始 COM AutoCaptions 对象。</param>
    internal WordAutoCaptions(MsWord.AutoCaptions autoCaptions)
    {
        _autoCaptions = autoCaptions ?? throw new ArgumentNullException(nameof(autoCaptions));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _autoCaptions != null ? new WordApplication(_autoCaptions.Application) : null;

    /// <inheritdoc/>
    public object Parent => _autoCaptions?.Parent;

    /// <inheritdoc/>
    public int Count => _autoCaptions?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordAutoCaption this[object index]
    {
        get
        {
            if (_autoCaptions == null) return null;
            try
            {
                var comAutoCaption = _autoCaptions[index];
                return comAutoCaption != null ? new WordAutoCaption(comAutoCaption) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 自动题注方法实现 (AutoCaptions Methods Implementation)

    /// <inheritdoc/>
    public void AutoInsert(string name, bool autoInsert, string captionLabel)
    {
        if (_autoCaptions == null || string.IsNullOrWhiteSpace(name)) return;

        try
        {
            // 尝试通过名称查找并设置
            var autoCap = this[name]; // 使用索引器
            if (autoCap != null)
            {
                autoCap.AutoInsert = autoInsert;
                autoCap.CaptionLabel = captionLabel ?? string.Empty;
            }
        }
        catch (COMException ex)
        {
            log.Error($"Failed to set AutoInsert for '{name}': {ex.Message}", ex);
        }
    }


    /// <inheritdoc/>
    public void CancelAutoInsert()
    {
        for (int i = 1; i <= this.Count; i++)
        {
            try
            {
                this[i].AutoInsert = false;
            }
            catch (COMException ex)
            {
                log.Error($"Failed to cancel auto insert for item {i}: {ex.Message}");
            }
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordAutoCaptions"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _autoCaptions != null)
        {
            Marshal.ReleaseComObject(_autoCaptions);
            _autoCaptions = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordAutoCaptions"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordAutoCaption> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordAutoCaption> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}