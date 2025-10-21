//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMergeFieldName : IWordMailMergeFieldName
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMergeFieldName));
    private MsWord.MailMergeFieldName? _fieldName;
    private bool _disposedValue;

    internal WordMailMergeFieldName(MsWord.MailMergeFieldName fieldName)
    {
        _fieldName = fieldName ?? throw new ArgumentNullException(nameof(fieldName));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _fieldName != null ? new WordApplication(_fieldName.Application) : null;

    public object? Parent => _fieldName?.Parent;

    public int Index => _fieldName?.Index ?? -1;

    public string? Name => _fieldName?.Name;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fieldName != null)
        {
            Marshal.ReleaseComObject(_fieldName);
            _fieldName = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}