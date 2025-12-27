//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Word 应用程序中所有任务窗格集合的封装实现类。
/// </summary>
internal class WordTaskPanes : IWordTaskPanes
{
    private MsWord.TaskPanes _taskPanes;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordTaskPanes"/> 类的新实例。
    /// </summary>
    /// <param name="taskPanes">要封装的原始 COM TaskPanes 对象。</param>
    internal WordTaskPanes(MsWord.TaskPanes taskPanes)
    {
        _taskPanes = taskPanes ?? throw new ArgumentNullException(nameof(taskPanes));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _taskPanes != null ? new WordApplication(_taskPanes.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _taskPanes?.Parent;

    /// <inheritdoc/>
    public int Count => _taskPanes?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordTaskPane this[WdTaskPanes index]
    {
        get
        {
            if (_taskPanes == null) return null;
            try
            {
                var comTaskPane = _taskPanes[(MsWord.WdTaskPanes)(int)index];
                return comTaskPane != null ? new WordTaskPane(comTaskPane) : null;
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to get TaskPane by index: {ex.Message}");
                return null;
            }
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTaskPanes"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _taskPanes != null)
        {
            Marshal.ReleaseComObject(_taskPanes);
            _taskPanes = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordTaskPanes"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordTaskPane> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordTaskPane> GetEnumerator()
    {
        foreach (var item in _taskPanes)
        {
            if (item is MsWord.TaskPane taskPane && taskPane != null)
                yield return new WordTaskPane(taskPane);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}