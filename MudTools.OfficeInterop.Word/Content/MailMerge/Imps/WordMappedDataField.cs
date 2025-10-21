//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;


internal class WordMappedDataField : IWordMappedDataField
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMappedDataField));
    private MsWord.MappedDataField? _mappedField;
    private bool _disposedValue;

    internal WordMappedDataField(MsWord.MappedDataField mappedField)
    {
        _mappedField = mappedField ?? throw new ArgumentNullException(nameof(mappedField));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _mappedField != null ? new WordApplication(_mappedField.Application) : null;

    public object? Parent => _mappedField?.Parent;

    public string? Name => _mappedField?.Name;

    public string? DataFieldName
    {
        get => _mappedField?.DataFieldName;
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mappedField != null)
        {
            Marshal.ReleaseComObject(_mappedField);
            _mappedField = null;
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