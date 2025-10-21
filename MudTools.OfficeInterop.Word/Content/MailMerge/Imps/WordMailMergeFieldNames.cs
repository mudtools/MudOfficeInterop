//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMergeFieldNames : IWordMailMergeFieldNames
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMergeFieldNames));
    private readonly DisposableList _disposableList = new();
    private MsWord.MailMergeFieldNames? _fieldNames;
    private bool _disposedValue;

    internal WordMailMergeFieldNames(MsWord.MailMergeFieldNames fieldNames)
    {
        _fieldNames = fieldNames ?? throw new ArgumentNullException(nameof(fieldNames));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _fieldNames != null ? new WordApplication(_fieldNames.Application) : null;

    public object? Parent => _fieldNames?.Parent;

    public int Count => _fieldNames?.Count ?? 0;

    public IWordMailMergeFieldName? this[int index]
    {
        get
        {
            if (_fieldNames == null || index < 1 || index > Count) return null;
            try
            {
                var comFieldName = _fieldNames[index];
                var result = comFieldName != null ? new WordMailMergeFieldName(comFieldName) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索 MailMergeFieldName 对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fieldNames != null)
        {
            Marshal.ReleaseComObject(_fieldNames);
            _disposableList.Dispose();
            _fieldNames = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordMailMergeFieldName> 实现

    public IEnumerator<IWordMailMergeFieldName> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}