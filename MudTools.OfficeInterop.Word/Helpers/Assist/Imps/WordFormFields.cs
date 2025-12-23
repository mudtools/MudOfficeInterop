//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表单域集合的封装实现类。
/// </summary>
internal class WordFormFields : IWordFormFields
{
    private MsWord.FormFields _formFields;
    private bool _disposedValue;

    internal WordFormFields(MsWord.FormFields formFields)
    {
        _formFields = formFields ?? throw new ArgumentNullException(nameof(formFields));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _formFields != null ? new WordApplication(_formFields.Application) : null;

    /// <inheritdoc/>
    public int Count => _formFields?.Count ?? 0;

    /// <inheritdoc/>
    public IWordFormField this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comField = _formFields[index];
            return new WordFormField(comField);
        }
    }

    /// <inheritdoc/>
    public IWordFormField this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            var comField = _formFields[name];
            return comField != null ? new WordFormField(comField) : null;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordFormField Add(IWordRange range, WdFieldType type)
    {
        if (range == null) throw new ArgumentNullException(nameof(range));

        try
        {
            var newField = _formFields.Add(((WordRange)range).InternalComObject, (MsWord.WdFieldType)(int)type);
            return new WordFormField(newField);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加表单域。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Contains(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return false;
        return _formFields[name] != null;
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_formFields == null) return;

        for (int i = Count; i >= 1; i--)
        {
            _formFields[i]?.Delete();
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _formFields != null)
        {
            Marshal.ReleaseComObject(_formFields);
            _formFields = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable 实现

    public IEnumerator<IWordFormField> GetEnumerator()
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