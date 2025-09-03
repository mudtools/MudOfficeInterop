//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示有权编辑文档特定部分的单个用户或用户组的封装实现类。
/// </summary>
internal class WordEditor : IWordEditor
{
    private MsWord.Editor _editor;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordEditor"/> 类的新实例。
    /// </summary>
    /// <param name="editor">要封装的原始 COM Editor 对象。</param>
    internal WordEditor(MsWord.Editor editor)
    {
        _editor = editor ?? throw new ArgumentNullException(nameof(editor));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _editor != null ? new WordApplication(_editor.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _editor?.Parent;

    /// <inheritdoc/>
    public string Name => _editor?.Name ?? string.Empty;

    /// <inheritdoc/>
    public string ID => _editor?.ID ?? string.Empty;

    /// <inheritdoc/>
    public IWordRange? Range => _editor?.Range != null ? new WordRange(_editor.Range) : null;

    /// <inheritdoc/>
    public IWordRange? NextRange => _editor?.NextRange != null ? new WordRange(_editor.NextRange) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _editor?.Delete();
    }

    /// <inheritdoc/>
    public void DeleteAll()
    {
        _editor?.DeleteAll();
    }

    /// <inheritdoc/>
    public void SelectAll()
    {
        _editor?.SelectAll();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordEditor"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _editor != null)
        {
            Marshal.ReleaseComObject(_editor);
            _editor = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordEditor"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}