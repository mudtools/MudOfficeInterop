//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档节实现类
/// </summary>
internal class WordSection : IWordSection
{
    private readonly MsWord.Section _section;
    private bool _disposedValue;
    private IWordRange? _range;
    private IWordPageSetup? _pageSetup;
    private IWordHeadersFooters? _wordFooters;
    private IWordHeadersFooters? _wordHeaders;
    private IWordBorders? _wordBorders;

    public IWordApplication? Application => _section != null ? new WordApplication(_section.Application) : null;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _section.Parent;

    /// <inheritdoc/>
    public int Index => _section?.Index ?? 0;

    public bool ProtectedForForms
    {
        get => _section != null && _section.ProtectedForForms;
        set
        {
            if (_section != null) _section.ProtectedForForms = value;
        }
    }

    /// <summary>
    /// 获取节范围
    /// </summary>
    public IWordRange? Range
    {
        get
        {
            if (_section == null) return null;
            _range ??= new WordRange(_section.Range);
            return _range;
        }
    }

    /// <inheritdoc/>
    public IWordHeadersFooters? Headers
    {
        get
        {
            if (_section == null) return null;
            _wordHeaders ??= new WordHeadersFooters(_section.Headers);
            return _wordHeaders;
        }
    }

    public IWordHeadersFooters? Footers
    {
        get
        {
            if (_section == null) return null;
            _wordFooters ??= new WordHeadersFooters(_section.Footers);
            return _wordFooters;
        }
    }


    /// <summary>
    /// 获取页面设置
    /// </summary>
    public IWordPageSetup? PageSetup
    {
        get
        {
            if (_section == null) return null;
            _pageSetup ??= new WordPageSetup(_section.PageSetup);
            return _pageSetup;
        }
    }

    public IWordBorders? Borders
    {
        get
        {
            if (_section == null) return null;
            _wordBorders ??= new WordBorders(_section.Borders);
            return _wordBorders;
        }
        set
        {
            if (_section == null || value == null)
                return;
            _wordBorders = value;
            _section.Borders = ((WordBorders)_wordBorders)._borders;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="section">COM Section 对象</param>
    internal WordSection(MsWord.Section section)
    {
        _section = section ?? throw new ArgumentNullException(nameof(section));
        _disposedValue = false;
    }

    /// <summary>
    /// 删除节
    /// </summary>
    public void Delete()
    {
        try
        {
            _section?.Range.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete section.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            Marshal.ReleaseComObject(_section);
            _range?.Dispose();
            _pageSetup?.Dispose();
            _wordFooters?.Dispose();
            _wordHeaders?.Dispose();
        }
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

