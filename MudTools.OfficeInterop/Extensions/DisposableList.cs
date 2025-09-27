//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using System.Reflection;

namespace MudTools.OfficeInterop;

internal class DisposableList : List<IDisposable>, IDisposable
{
    private bool _disposed = false;
    private readonly object _lockObject = new();

    /// <summary>
    /// 用于记录日志的静态日志记录器。
    /// </summary>
    private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

    /// <summary>
    /// 添加一个可释放对象到列表
    /// </summary>
    public new void Add(IDisposable item)
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(DisposableList));
        }

        lock (_lockObject)
        {
            base.Add(item);
        }
    }

    /// <summary>
    /// 添加多个可释放对象到列表
    /// </summary>
    public new void AddRange(IEnumerable<IDisposable> items)
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(DisposableList));

        if (items == null)
            throw new ArgumentNullException(nameof(items));

        lock (_lockObject)
        {
            base.AddRange(items);
        }
    }

    /// <summary>
    /// 尝试移除并释放指定的对象
    /// </summary>
    public bool RemoveAndDispose(IDisposable item)
    {
        if (_disposed)
            return false;

        lock (_lockObject)
        {
            var removed = base.Remove(item);
            if (removed)
            {
                SafeDispose(item);
            }
            return removed;
        }
    }

    /// <summary>
    /// 释放所有对象并清空列表
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            List<IDisposable> itemsToDispose;
            lock (_lockObject)
            {
                itemsToDispose = this.ToList();
                base.Clear();
            }

            List<Exception> exceptions = new List<Exception>();

            // 在锁外释放对象，避免死锁和性能问题
            foreach (var item in itemsToDispose)
            {
                try
                {
                    item?.Dispose();
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                    _log.Error($"Error disposing item of type {item?.GetType().Name}", ex);
                }
            }

            // 如果有异常，记录日志但不抛出异常（Dispose方法不应抛出异常）
            if (exceptions.Count > 0)
            {
                _log.Error($"One or more errors occurred while disposing {exceptions.Count} items.");
            }
        }

        _disposed = true;
    }

    /// <summary>
    /// 安全释放单个对象（不抛出异常）
    /// </summary>
    private void SafeDispose(IDisposable item)
    {
        try
        {
            item?.Dispose();
        }
        catch (Exception ex)
        {
            _log.Error($"An error occurred while disposing an item of type {item?.GetType().Name}", ex);
        }
    }

    /// <summary>
    /// 获取是否已释放
    /// </summary>
    public bool IsDisposed => _disposed;

    ~DisposableList()
    {
        Dispose(false);
    }
}