//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordTableOfContents : IWordTableOfContents
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTableOfContents));
    private MsWord.TableOfContents? _tableOfContents;
    private bool _disposedValue;

    internal WordTableOfContents(MsWord.TableOfContents tableOfContents)
    {
        _tableOfContents = tableOfContents ?? throw new ArgumentNullException(nameof(tableOfContents));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _tableOfContents != null ? new WordApplication(_tableOfContents.Application) : null;

    public object? Parent => _tableOfContents?.Parent;

    public IWordRange? Range => _tableOfContents?.Range != null ? new WordRange(_tableOfContents.Range) : null;

    public IWordHeadingStyles? HeadingStyles => _tableOfContents != null ? new WordHeadingStyles(_tableOfContents.HeadingStyles) : null;


    public string TableID
    {
        get => _tableOfContents?.TableID ?? string.Empty;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.TableID = value;
        }
    }

    public bool UseHyperlinks
    {
        get => _tableOfContents?.UseHyperlinks ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.UseHyperlinks = value;
        }
    }

    public int LowerHeadingLevel
    {
        get => _tableOfContents?.LowerHeadingLevel ?? 0;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.LowerHeadingLevel = value;
        }
    }

    public int UpperHeadingLevel
    {
        get => _tableOfContents?.UpperHeadingLevel ?? 0;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.UpperHeadingLevel = value;
        }
    }

    public bool UseFields
    {
        get => _tableOfContents?.UseFields ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.UseFields = value;
        }
    }

    public bool UseHeadingStyles
    {
        get => _tableOfContents?.UseHeadingStyles ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.UseHeadingStyles = value;
        }
    }

    public bool RightAlignPageNumbers
    {
        get => _tableOfContents?.RightAlignPageNumbers ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.RightAlignPageNumbers = value;
        }
    }

    public bool IncludePageNumbers
    {
        get => _tableOfContents?.IncludePageNumbers ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.IncludePageNumbers = value;
        }
    }

    public bool HidePageNumbersInWeb
    {
        get => _tableOfContents?.HidePageNumbersInWeb ?? false;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.HidePageNumbersInWeb = value;
        }
    }

    public WdTabLeader TabLeader
    {
        get => _tableOfContents != null ? _tableOfContents.TabLeader.EnumConvert(WdTabLeader.wdTabLeaderLines) : WdTabLeader.wdTabLeaderLines;
        set
        {
            if (_tableOfContents != null)
                _tableOfContents.TabLeader = value.EnumConvert(MsWord.WdTabLeader.wdTabLeaderLines);
        }
    }



    #endregion

    #region 方法实现

    public void Update()
    {
        try
        {
            _tableOfContents?.Update();
        }
        catch (Exception ex)
        {
            log.Error("更新目录失败。", ex);
            throw new InvalidOperationException("更新目录失败。", ex);
        }
    }

    public void UpdatePageNumbers()
    {
        try
        {
            _tableOfContents?.UpdatePageNumbers();
        }
        catch (Exception ex)
        {
            log.Error("更新目录页码失败。", ex);
            throw new InvalidOperationException("更新目录页码失败。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _tableOfContents?.Delete();
        }
        catch (Exception ex)
        {
            log.Error("删除目录失败。", ex);
            throw new InvalidOperationException("删除目录失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tableOfContents != null)
        {
            Marshal.ReleaseComObject(_tableOfContents);
            _tableOfContents = null;
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