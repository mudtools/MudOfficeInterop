//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表示下拉列表或组合框内容控件中一个列表项的封装实现类。
/// </summary>
internal class WordContentControlListEntry : IWordContentControlListEntry
{
    private MsWord.ContentControlListEntry _listEntry;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordContentControlListEntry"/> 类的新实例。
    /// </summary>
    /// <param name="listEntry">要封装的原始 COM ContentControlListEntry 对象。</param>
    internal WordContentControlListEntry(MsWord.ContentControlListEntry listEntry)
    {
        _listEntry = listEntry ?? throw new ArgumentNullException(nameof(listEntry));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listEntry != null ? new WordApplication(_listEntry.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _listEntry?.Parent;

    /// <inheritdoc/>
    public int Index
    {
        get => _listEntry?.Index ?? 0;
        set { if (_listEntry != null) _listEntry.Index = value; }
    }

    /// <inheritdoc/>
    public string Text
    {
        get => _listEntry?.Text ?? string.Empty;
        set { if (_listEntry != null) _listEntry.Text = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string Value
    {
        get => _listEntry?.Value ?? string.Empty;
        set { if (_listEntry != null) _listEntry.Value = value ?? string.Empty; }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _listEntry?.Delete();
    }

    /// <inheritdoc/>
    public void MoveDown()
    {
        _listEntry?.MoveDown();
    }

    /// <inheritdoc/>
    public void MoveUp()
    {
        _listEntry?.MoveUp();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _listEntry?.Select();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordContentControlListEntry"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listEntry != null)
        {
            Marshal.ReleaseComObject(_listEntry);
            _listEntry = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordContentControlListEntry"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}