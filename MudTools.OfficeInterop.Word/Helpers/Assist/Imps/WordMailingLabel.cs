//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示邮件标签的全局邮件标签首选项的封装实现类。
/// </summary>
internal class WordMailingLabel : IWordMailingLabel
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailingLabel));
    private MsWord.MailingLabel _mailingLabel;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordMailingLabel"/> 类的新实例。
    /// </summary>
    /// <param name="mailingLabel">要封装的原始 COM MailingLabel 对象。</param>
    internal WordMailingLabel(MsWord.MailingLabel mailingLabel)
    {
        _mailingLabel = mailingLabel ?? throw new ArgumentNullException(nameof(mailingLabel));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _mailingLabel != null ? new WordApplication(_mailingLabel.Application) : null;

    /// <inheritdoc/>
    public object Parent => _mailingLabel?.Parent;

    /// <inheritdoc/>
    public int Creator => _mailingLabel?.Creator ?? 0;

    #endregion

    #region 邮件标签属性实现 (Mailing Label Properties Implementation)

    /// <inheritdoc/>
    public string DefaultLabelName
    {
        get => _mailingLabel?.DefaultLabelName ?? string.Empty;
        set { if (_mailingLabel != null) _mailingLabel.DefaultLabelName = value ?? string.Empty; }
    }

    #endregion

    #region 邮件标签方法实现 (Mailing Label Methods Implementation)

    public IWordDocument CreateNewDocument(string? name = null, string? address = null,
        string? autoText = null, bool? extractAddress = null,
        WdPaperTray laserTray = WdPaperTray.wdPrinterDefaultBin,
        bool? printEPostageLabel = null, object? vertical = null)
    {
        if (_mailingLabel == null) return null;
        try
        {
            var document = _mailingLabel.CreateNewDocument(name.ComArgsVal(),
                address.ComArgsVal(), autoText.ComArgsVal(),
                extractAddress.ComArgsVal(), laserTray, printEPostageLabel.ComArgsVal(),
                vertical != null ? vertical : Type.Missing);
            return document != null ? new WordDocument(document) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to create label: {ex.Message}");
            return null;
        }
    }

    /// <inheritdoc/>
    public void PrintOut(string? name, string? address, string? extractAddress,
         WdPaperTray laserTray = WdPaperTray.wdPrinterDefaultBin)
    {
        _mailingLabel?.PrintOut(
            name.ComArgsVal(),
            address.ComArgsVal(),
            extractAddress.ComArgsVal(), laserTray);
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordMailingLabel"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mailingLabel != null)
        {
            Marshal.ReleaseComObject(_mailingLabel);
            _mailingLabel = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordMailingLabel"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}