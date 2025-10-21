//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMergeFields : IWordMailMergeFields
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMergeFields));
    private readonly DisposableList _disposableList = new();
    private MsWord.MailMergeFields? _mailMergeFields;
    private bool _disposedValue;

    internal WordMailMergeFields(MsWord.MailMergeFields mailMergeFields)
    {
        _mailMergeFields = mailMergeFields ?? throw new ArgumentNullException(nameof(mailMergeFields));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _mailMergeFields != null ? new WordApplication(_mailMergeFields.Application) : null;

    public object? Parent => _mailMergeFields?.Parent;

    public int Count => _mailMergeFields?.Count ?? 0;

    public IWordMailMergeField? this[int index]
    {
        get
        {
            if (_mailMergeFields == null || index < 1 || index > Count) return null;
            try
            {
                var comField = _mailMergeFields[index];
                var result = comField != null ? new WordMailMergeField(comField) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索 MailMergeField 对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    public IWordMailMergeField Add(IWordRange range, string fieldName)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentNullException(nameof(fieldName));

        if (_mailMergeFields == null)
            throw new InvalidOperationException("邮件合并域集合不可用。");

        try
        {
            var comRange = (range as WordRange)?._range;
            if (comRange == null)
                throw new ArgumentException("提供的范围对象无效。", nameof(range));

            var comField = _mailMergeFields.Add(comRange, fieldName);
            var fieldWrapper = new WordMailMergeField(comField);
            _disposableList.Add(fieldWrapper);
            return fieldWrapper;
        }
        catch (Exception ex)
        {
            log.Error("向文档添加新邮件合并域失败。", ex);
            throw new InvalidOperationException("向文档添加新邮件合并域失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mailMergeFields != null)
        {
            Marshal.ReleaseComObject(_mailMergeFields);
            _disposableList.Dispose();
            _mailMergeFields = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordMailMergeField> 实现

    public IEnumerator<IWordMailMergeField> GetEnumerator()
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