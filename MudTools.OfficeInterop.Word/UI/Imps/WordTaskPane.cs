//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Word 应用程序中一个任务窗格的封装实现类。
/// </summary>
internal class WordTaskPane : IWordTaskPane
{
    private MsWord.TaskPane _taskPane;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordTaskPane"/> 类的新实例。
    /// </summary>
    /// <param name="taskPane">要封装的原始 COM TaskPane 对象。</param>
    internal WordTaskPane(MsWord.TaskPane taskPane)
    {
        _taskPane = taskPane ?? throw new ArgumentNullException(nameof(taskPane));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _taskPane != null ? new WordApplication(_taskPane.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _taskPane?.Parent;

    /// <inheritdoc/>
    public int Creator => _taskPane?.Creator ?? 0;

    #endregion

    #region 任务窗格属性实现 (Task Pane Properties Implementation) 

    /// <inheritdoc/>
    public bool Visible
    {
        get => _taskPane?.Visible ?? false;
        set
        {
            if (_taskPane != null)
                _taskPane.Visible = value;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTaskPane"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _taskPane != null)
        {
            Marshal.ReleaseComObject(_taskPane);
            _taskPane = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordTaskPane"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}