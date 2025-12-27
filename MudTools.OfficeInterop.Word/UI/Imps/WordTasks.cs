//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示当前在系统上运行的所有应用程序任务集合的封装实现类。
/// </summary>
internal class WordTasks : IWordTasks
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTasks));
    private MsWord.Tasks _tasks;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordTasks"/> 类的新实例。
    /// </summary>
    /// <param name="tasks">要封装的原始 COM Tasks 对象。</param>
    internal WordTasks(MsWord.Tasks tasks)
    {
        _tasks = tasks ?? throw new ArgumentNullException(nameof(tasks));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _tasks != null ? new WordApplication(_tasks.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _tasks?.Parent;

    /// <inheritdoc/>
    public int Count => _tasks?.Count ?? 0;

    /// <inheritdoc/>
    public IWordTask this[object index]
    {
        get
        {
            if (_tasks == null) return null;
            try
            {
                var comTask = _tasks[index];
                return comTask != null ? new WordTask(comTask) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public bool Exists(string name)
    {
        return _tasks?.Exists(name) ?? false;
    }

    /// <inheritdoc/>
    public void ExitWindows()
    {
        _tasks?.ExitWindows();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTasks"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tasks != null)
        {
            Marshal.ReleaseComObject(_tasks);
            _tasks = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordTasks"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordTask> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordTask> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}