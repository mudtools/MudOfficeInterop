//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

// <summary>
/// 对 <see cref="AutoCorrectEntry"/> COM 对象的封装实现类。
/// 提供安全的属性访问和资源释放机制。
/// </summary>
internal class WordAutoCorrectEntry : IWordAutoCorrectEntry
{
    private MsWord.AutoCorrectEntry _entry;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM AutoCorrectEntry 对象初始化封装实例。
    /// </summary>
    /// <param name="entry">原始的 AutoCorrectEntry COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 entry 为 null 时抛出。</exception>
    internal WordAutoCorrectEntry(MsWord.AutoCorrectEntry entry)
    {
        _entry = entry ?? throw new ArgumentNullException(nameof(entry));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public string Name
    {
        get
        {
            return _entry?.Name ?? string.Empty;
        }
    }

    /// <inheritdoc />
    public string Value
    {
        get
        {
            return _entry?.Value ?? string.Empty;
        }
        set
        {
            value ??= string.Empty;
            _entry.Value = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public void Delete()
    {
        if (_disposedValue || _entry == null) return;

        try
        {
            _entry.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("删除自动更正条目失败。", ex);
        }
    }

    public void Apply(IWordRange range)
    {
        if (range == null) return;
        try
        {
            _entry.Apply(((WordRange)range).InternalComObject);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("应用自动更正条目失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源（COM 对象）。
    /// </summary>
    /// <param name="disposing">为 true 表示由 Dispose() 显式调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _entry != null)
        {
            try
            {
                Marshal.ReleaseComObject(_entry);
            }
            catch (InvalidComObjectException)
            {
                // 忽略已释放对象
            }
            finally
            {
                _entry = null;
            }
        }

        _disposedValue = true;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}