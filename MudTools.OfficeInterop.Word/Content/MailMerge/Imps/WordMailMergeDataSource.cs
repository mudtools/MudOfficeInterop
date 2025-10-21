//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMergeDataSource : IWordMailMergeDataSource
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMergeDataSource));
    private MsWord.MailMergeDataSource? _dataSource;
    private bool _disposedValue;

    internal WordMailMergeDataSource(MsWord.MailMergeDataSource dataSource)
    {
        _dataSource = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _dataSource != null ? new WordApplication(_dataSource.Application) : null;

    public object? Parent => _dataSource?.Parent;

    public int FirstRecord
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.FirstRecord;
        }
        set
        {
            if (_dataSource != null)
                _dataSource.FirstRecord = value;
        }
    }

    public int LastRecord
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.LastRecord;
        }
        set
        {
            if (_dataSource != null)
                _dataSource.LastRecord = value;
        }
    }

    public string TableName => _dataSource?.TableName ?? string.Empty;

    public WdMailMergeActiveRecord ActiveRecord
    {
        get => _dataSource?.ActiveRecord.EnumConvert(WdMailMergeActiveRecord.wdNoActiveRecord) ?? WdMailMergeActiveRecord.wdNoActiveRecord;
        set
        {
            if (_dataSource != null)
                _dataSource.ActiveRecord = value.EnumConvert(MsWord.WdMailMergeActiveRecord.wdNoActiveRecord);
        }
    }

    public WdMailMergeDataSource HeaderSourceType
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.HeaderSourceType.EnumConvert(WdMailMergeDataSource.wdNoMergeInfo);
        }
    }

    public WdMailMergeDataSource Type
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.Type.EnumConvert(WdMailMergeDataSource.wdNoMergeInfo);
        }
    }

    public bool InvalidAddress
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.InvalidAddress;
        }
        set
        {
            if (_dataSource != null)
                _dataSource.InvalidAddress = value;
        }
    }

    public string HeaderSourceName
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.HeaderSourceName;
        }

    }

    public string Name
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.Name;
        }
    }

    public bool Included
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.Included;
        }
        set
        {
            if (_dataSource != null)
                _dataSource.Included = value;
        }
    }

    public int RecordCount
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.RecordCount;
        }
    }

    public string? ConnectString => _dataSource?.ConnectString;

    public string? QueryString
    {
        get
        {
            if (_dataSource == null)
                throw new InvalidOperationException("数据源不可用。");
            return _dataSource.QueryString;
        }
        set
        {
            if (_dataSource != null)
                _dataSource.QueryString = value;
        }
    }

    public IWordMailMergeFieldNames? FieldNames => _dataSource?.FieldNames != null ? new WordMailMergeFieldNames(_dataSource.FieldNames) : null;

    public IWordMailMergeDataFields? DataFields => _dataSource?.DataFields != null ? new WordMailMergeDataFields(_dataSource.DataFields) : null;

    public IWordMappedDataFields? MappedDataFields => _dataSource?.MappedDataFields != null ? new WordMappedDataFields(_dataSource.MappedDataFields) : null;
    #endregion

    #region 方法实现

    public void SetAllIncludedFlags(bool Included)
    {
        try
        {
            _dataSource?.SetAllIncludedFlags(Included);
        }
        catch (Exception ex)
        {
            log.Error($"设置所有记录的 Included 标志为 {Included} 失败。", ex);
            throw new InvalidOperationException($"设置所有记录的 Included 标志为 {Included} 失败。", ex);
        }
    }

    public void SetAllErrorFlags(bool invalid, string invalidComment)
    {
        try
        {
            _dataSource?.SetAllErrorFlags(invalid, invalidComment);
        }
        catch (Exception ex)
        {
            log.Error($"设置所有记录的错误标志为 {invalid}，错误信息为 '{invalidComment}' 失败。", ex);
            throw new InvalidOperationException($"设置所有记录的错误标志为 {invalid}，错误信息为 '{invalidComment}' 失败。", ex);
        }
    }

    public void Close()
    {
        if (_dataSource == null)
            throw new InvalidOperationException("数据源不可用。");

        try
        {
            _dataSource.Close();
        }
        catch (Exception ex)
        {
            log.Error("关闭数据源失败。", ex);
            throw new InvalidOperationException("关闭数据源失败。", ex);
        }
    }

    public bool FindRecord(string text, string? fieldName = null)
    {
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentNullException(nameof(fieldName));
        if (string.IsNullOrEmpty(text))
            throw new ArgumentNullException(nameof(text));

        if (_dataSource == null)
            throw new InvalidOperationException("数据源不可用。");

        try
        {
            return _dataSource.FindRecord(text, fieldName.ComArgsVal());
        }
        catch (Exception ex)
        {
            log.Error($"在字段 '{fieldName}' 中查找文本 '{text}' 失败。", ex);
            throw new InvalidOperationException($"在字段 '{fieldName}' 中查找文本 '{text}' 失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dataSource != null)
        {
            Marshal.ReleaseComObject(_dataSource);
            _dataSource = null;
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