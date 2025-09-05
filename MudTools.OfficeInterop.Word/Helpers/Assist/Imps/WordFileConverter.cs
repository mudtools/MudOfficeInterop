//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示用于打开或保存文件的文件转换器的封装实现类。
/// </summary>
internal class WordFileConverter : IWordFileConverter
{
    private MsWord.FileConverter _fileConverter;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordFileConverter"/> 类的新实例。
    /// </summary>
    /// <param name="fileConverter">要封装的原始 COM FileConverter 对象。</param>
    internal WordFileConverter(MsWord.FileConverter fileConverter)
    {
        _fileConverter = fileConverter ?? throw new ArgumentNullException(nameof(fileConverter));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _fileConverter != null ? new WordApplication(_fileConverter.Application) : null;

    /// <inheritdoc/>
    public object Parent => _fileConverter?.Parent;

    /// <inheritdoc/>
    public bool CanOpen => _fileConverter?.CanOpen ?? false;

    /// <inheritdoc/>
    public bool CanSave => _fileConverter?.CanSave ?? false;

    /// <inheritdoc/>
    public string ClassName => _fileConverter?.ClassName ?? string.Empty;

    /// <inheritdoc/>
    public int Creator => _fileConverter?.Creator ?? 0;

    /// <inheritdoc/>
    public string Extensions => _fileConverter?.Extensions ?? string.Empty;

    /// <inheritdoc/>
    public string FormatName => _fileConverter?.FormatName ?? string.Empty;

    /// <inheritdoc/>
    public string Name
    {
        get => _fileConverter?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public int OpenFormat => _fileConverter?.OpenFormat ?? 0;

    /// <inheritdoc/>
    public int SaveFormat => _fileConverter?.SaveFormat ?? 0;

    /// <inheritdoc/>
    public string Path => _fileConverter?.Path ?? string.Empty;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordFileConverter"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fileConverter != null)
        {
            Marshal.ReleaseComObject(_fileConverter);
            _fileConverter = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordFileConverter"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
