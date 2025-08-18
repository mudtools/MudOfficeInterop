//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 超链接实现类
/// </summary>
internal class WordHyperlink : IWordHyperlink
{
    private readonly MsWord.Hyperlink _hyperlink;
    private bool _disposedValue;

    public string Address
    {
        get => _hyperlink.Address;
        set => _hyperlink.Address = value;
    }

    public string TextToDisplay
    {
        get => _hyperlink.TextToDisplay;
        set => _hyperlink.TextToDisplay = value;
    }

    public object Parent => _hyperlink.Parent;

    internal WordHyperlink(MsWord.Hyperlink hyperlink)
    {
        _hyperlink = hyperlink ?? throw new ArgumentNullException(nameof(hyperlink));
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _hyperlink.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete hyperlink.", ex);
        }
    }

    public void Follow()
    {
        try
        {
            _hyperlink.Follow();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to follow hyperlink.", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}