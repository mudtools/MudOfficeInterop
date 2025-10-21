
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;


internal class WordMappedDataFields : IWordMappedDataFields
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMappedDataFields));
    private readonly DisposableList _disposableList = new();
    private MsWord.MappedDataFields? _mappedFields;
    private bool _disposedValue;

    internal WordMappedDataFields(MsWord.MappedDataFields mappedFields)
    {
        _mappedFields = mappedFields ?? throw new ArgumentNullException(nameof(mappedFields));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _mappedFields != null ? new WordApplication(_mappedFields.Application) : null;

    public object? Parent => _mappedFields?.Parent;

    public int Count => _mappedFields?.Count ?? 0;

    public IWordMappedDataField? this[WdMappedDataFields index]
    {
        get
        {
            if (_mappedFields == null) return null;
            try
            {
                var comMappedField = _mappedFields[index.EnumConvert(MsWord.WdMappedDataFields.wdFirstName)];
                var result = comMappedField != null ? new WordMappedDataField(comMappedField) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索 MappedDataField 对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mappedFields != null)
        {
            Marshal.ReleaseComObject(_mappedFields);
            _disposableList.Dispose();
            _mappedFields = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordMappedDataField> 实现

    public IEnumerator<IWordMappedDataField> GetEnumerator()
    {
        if (_mappedFields == null)
            yield break;

        foreach (var item in _mappedFields)
        {
            if (item != null && item is MsWord.MappedDataField mappedField)
            {
                var result = new WordMappedDataField(mappedField);
                _disposableList.Add(result);
                yield return result;
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}