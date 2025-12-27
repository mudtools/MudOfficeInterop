//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示一个自动图文集条目的封装实现类。
/// </summary>
internal class WordAutoCaption : IWordAutoCaption
{
    private MsWord.AutoCaption _autoCaption;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordAutoCaption"/> 类的新实例。
    /// </summary>
    /// <param name="autoCaption">要封装的原始 COM AutoCaption 对象。</param>
    internal WordAutoCaption(MsWord.AutoCaption autoCaption)
    {
        _autoCaption = autoCaption ?? throw new ArgumentNullException(nameof(autoCaption));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication? Application => _autoCaption != null ? new WordApplication(_autoCaption.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _autoCaption?.Parent;

    /// <inheritdoc/>
    public int Creator => _autoCaption?.Creator ?? 0;

    /// <inheritdoc/>
    public int Index => _autoCaption?.Index ?? 0;

    #endregion

    #region 自动题注属性实现 (AutoCaption Properties Implementation)

    /// <inheritdoc/>
    public bool AutoInsert
    {
        get => _autoCaption?.AutoInsert ?? false;
        set
        {
            if (_autoCaption != null)
                _autoCaption.AutoInsert = value;
        }
    }

    /// <inheritdoc/>
    public string CaptionLabel
    {
        get => _autoCaption?.get_CaptionLabel()?.ToString() ?? string.Empty;
        set
        {
            if (_autoCaption != null)
                _autoCaption.set_CaptionLabel(value ?? string.Empty);
        }
    }

    /// <inheritdoc/>
    public string Name => _autoCaption?.Name ?? string.Empty;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordAutoCaption"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _autoCaption != null)
        {
            Marshal.ReleaseComObject(_autoCaption);
            _autoCaption = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordAutoCaption"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}