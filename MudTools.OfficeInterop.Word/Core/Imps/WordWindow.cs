//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档窗口实现类
/// </summary>
internal class WordWindow : IWordWindow
{
    private readonly MsWord.Window _window;
    private bool _disposedValue;

    public string Caption => _window.Caption;

    public int Height
    {
        get => _window.Height;
        set => _window.Height = value;
    }

    public int Width
    {
        get => _window.Width;
        set => _window.Width = value;
    }

    public int Left
    {
        get => _window.Left;
        set => _window.Left = value;
    }

    public int Top
    {
        get => _window.Top;
        set => _window.Top = value;
    }

    public object Parent => _window.Parent;

    /// <summary>
    /// 获取Excel应用程序窗口的句柄
    /// </summary>
    public int? Hwnd => _window?.Hwnd;

    internal WordWindow(MsWord.Window window)
    {
        _window = window ?? throw new ArgumentNullException(nameof(window));
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _window.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate window.", ex);
        }
    }

    public void Close()
    {
        try
        {
            _window.Close();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close window.", ex);
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