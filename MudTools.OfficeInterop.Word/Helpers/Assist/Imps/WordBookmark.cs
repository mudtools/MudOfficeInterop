//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 书签实现类
/// </summary>
internal class WordBookmark : IWordBookmark
{
    private readonly MsWord.Bookmark _bookmark;
    private bool _disposedValue;
    private IWordRange _range;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _bookmark != null ? new WordApplication(_bookmark.Application) : null;

    public string Name => _bookmark.Name;

    public IWordRange Range
    {
        get
        {
            if (_range == null)
            {
                _range = new WordRange(_bookmark.Range);
            }
            return _range;
        }
    }

    public object? Parent => _bookmark.Parent;

    internal WordBookmark(MsWord.Bookmark bookmark)
    {
        _bookmark = bookmark ?? throw new ArgumentNullException(nameof(bookmark));
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _bookmark.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete bookmark.", ex);
        }
    }

    public void Select()
    {
        try
        {
            _bookmark.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select bookmark.", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _range?.Dispose();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}