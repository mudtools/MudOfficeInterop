using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// <see cref="IWordAutoTextEntry"/> 接口的具体实现类，封装一个 Word AutoTextEntry COM 对象。
/// 管理 COM 生命周期，防止资源泄漏。
/// </summary>
internal class WordAutoTextEntry : IWordAutoTextEntry
{
    private AutoTextEntry _entry;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数：包装一个现有的 AutoTextEntry COM 对象
    /// </summary>
    /// <param name="entry">Interop 中的 AutoTextEntry 实例</param>
    internal WordAutoTextEntry(AutoTextEntry entry)
    {
        _entry = entry ?? throw new ArgumentNullException(nameof(entry));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取自动图文集条目的名称（例如："SignatureBlock"）
    /// </summary>
    public string Name => _entry?.Name;

    /// <summary>
    /// 获取或设置自动图文集条目的内容值。
    /// 注意：读取时返回纯文本（不包含格式），写入时将替换整个内容。
    /// </summary>
    public string Value
    {
        get
        {
            return _entry.Value;
        }
        set
        {
            _entry.Value = value;
        }
    }

    /// <summary>
    /// 获取此自动图文集条目样式名
    /// </summary>
    public string StyleName => _entry.StyleName;

    #endregion

    #region 方法实现
    /// <summary>
    /// 从模板中删除此自动图文集条目
    /// </summary>
    public void Delete()
    {
        if (_disposedValue || _entry == null) return;

        try
        {
            _entry.Delete();
        }
        catch (COMException)
        {
            // 忽略删除错误（可能已被删除）
        }
        finally
        {
            Dispose();
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _entry != null)
        {
            Marshal.ReleaseComObject(_entry);
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}