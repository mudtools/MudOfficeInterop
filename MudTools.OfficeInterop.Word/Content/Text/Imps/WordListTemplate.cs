//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ListTemplate 的封装实现类。
/// </summary>
internal class WordListTemplate : IWordListTemplate
{
    internal MsWord.ListTemplate _listTemplate;
    private bool _disposedValue;

    internal WordListTemplate(MsWord.ListTemplate listTemplate)
    {
        _listTemplate = listTemplate ?? throw new ArgumentNullException(nameof(listTemplate));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _listTemplate != null ? new WordApplication(_listTemplate.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _listTemplate?.Parent;

    /// <inheritdoc/>
    public string Name => _listTemplate?.Name ?? string.Empty;

    /// <inheritdoc/>
    public bool OutlineNumbered
    {
        get => _listTemplate?.OutlineNumbered ?? false;
        set
        {
            if (_listTemplate != null)
                _listTemplate.OutlineNumbered = value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordListLevels ListLevels => _listTemplate?.ListLevels != null ? new WordListLevels(_listTemplate.ListLevels) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordListTemplate Convert(IWordListLevel listLevel)
    {
        if (_listTemplate != null)
        {
            MsWord.ListTemplate listTemplate;
            if (listLevel is WordListLevel wordListLevel && wordListLevel != null)
            {
                listTemplate = _listTemplate.Convert(wordListLevel._listLevel);
            }
            else
            {
                listTemplate = _listTemplate.Convert();
            }

            return new WordListTemplate(listTemplate);
        }
        return null;
    }

    /// <inheritdoc/>
    public IWordListLevel GetListLevel(int level)
    {
        if (_listTemplate != null && level >= 1 && level <= 9)
        {
            var listLevel = _listTemplate.ListLevels[level];
            return new WordListLevel(listLevel);
        }
        return null;
    }

    /// <inheritdoc/>
    public void SetAllLevelNumberFormats(string[] formats)
    {
        if (_listTemplate != null && formats != null)
        {
            for (int i = 0; i < formats.Length && i < 9; i++)
            {
                var listLevel = _listTemplate.ListLevels[i + 1];
                if (listLevel != null)
                {
                    listLevel.NumberFormat = formats[i];
                }
            }
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            (ListLevels as IDisposable)?.Dispose();

            if (_listTemplate != null)
            {
                Marshal.ReleaseComObject(_listTemplate);
            }
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
