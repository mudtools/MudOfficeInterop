//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示当前在系统上运行的应用程序任务的封装实现类。
/// </summary>
internal class WordTask : IWordTask
{
    private MsWord.Task _task;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordTask"/> 类的新实例。
    /// </summary>
    /// <param name="task">要封装的原始 COM Task 对象。</param>
    internal WordTask(MsWord.Task task)
    {
        _task = task ?? throw new ArgumentNullException(nameof(task));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _task != null ? new WordApplication(_task.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _task?.Parent;

    /// <inheritdoc/>
    public int Creator => _task?.Creator ?? 0;

    /// <inheritdoc/>
    public int Left
    {
        get => _task?.Left ?? 0;
        set { if (_task != null) _task.Left = value; }
    }

    /// <inheritdoc/>
    public int Top
    {
        get => _task?.Top ?? 0;
        set { if (_task != null) _task.Top = value; }
    }

    /// <inheritdoc/>
    public int Width
    {
        get => _task?.Width ?? 0;
        set { if (_task != null) _task.Width = value; }
    }

    /// <inheritdoc/>
    public int Height
    {
        get => _task?.Height ?? 0;
        set { if (_task != null) _task.Height = value; }
    }

    /// <inheritdoc/>
    public WdWindowState WindowState
    {
        get => _task?.WindowState != null ? (WdWindowState)(int)_task?.WindowState : WdWindowState.wdWindowStateNormal;
        set
        {
            if (_task != null) _task.WindowState = (MsWord.WdWindowState)(int)value;
        }
    }

    /// <inheritdoc/>
    public string Name => _task?.Name ?? string.Empty;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Activate()
    {
        _task?.Activate();
    }

    /// <inheritdoc/>
    public void Move(int Left, int Top)
    {
        _task?.Move(Left, Top);
    }

    /// <inheritdoc/>
    public void Resize(int Width, int Height)
    {
        _task?.Resize(Width, Height);
    }

    /// <inheritdoc/>
    public void Close()
    {
        _task?.Close();
    }

    /// <inheritdoc/>
    public void SendWindowMessage(int Message, int wParam, int lParam)
    {
        _task?.SendWindowMessage(Message, wParam, lParam);
    }


    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTask"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _task != null)
        {
            Marshal.ReleaseComObject(_task);
            _task = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordTask"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}