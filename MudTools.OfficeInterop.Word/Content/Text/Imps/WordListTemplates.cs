//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ListTemplates 的封装实现类。
/// </summary>
internal class WordListTemplates : IWordListTemplates
{
    private MsWord.ListTemplates _listTemplates;
    private bool _disposedValue;

    internal WordListTemplates(MsWord.ListTemplates listTemplates)
    {
        _listTemplates = listTemplates ?? throw new ArgumentNullException(nameof(listTemplates));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listTemplates != null ? new WordApplication(_listTemplates.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _listTemplates?.Parent;

    /// <inheritdoc/>
    public int Count => _listTemplates?.Count ?? 0;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordListTemplate this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comListTemplate = _listTemplates[index];
                return new WordListTemplate(comListTemplate);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IWordListTemplate this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            try
            {
                var comListTemplate = _listTemplates[name];
                return comListTemplate != null ? new WordListTemplate(comListTemplate) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordListTemplate Add(bool outlineNumbered, bool builtIn)
    {
        try
        {
            var newListTemplate = _listTemplates.Add(outlineNumbered, builtIn);
            return new WordListTemplate(newListTemplate);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加列表模板。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Contains(string name)
    {
        if (_disposedValue || string.IsNullOrWhiteSpace(name)) return false;

        try
        {
            return _listTemplates[name] != null;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public List<string> GetNames()
    {
        var names = new List<string>();
        for (int i = 1; i <= Count; i++)
        {
            var template = _listTemplates[i];
            if (template?.Name != null)
                names.Add(template.Name.ToString());
        }
        return names;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordListTemplate> GetEnumerator()
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

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listTemplates != null)
        {
            Marshal.ReleaseComObject(_listTemplates);
            _listTemplates = null;
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