//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表示有权编辑文档特定部分的用户或用户组集合的封装实现类。
/// </summary>
internal class WordEditors : IWordEditors
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordEditors));

    private MsWord.Editors _editors;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordEditors"/> 类的新实例。
    /// </summary>
    /// <param name="editors">要封装的原始 COM Editors 对象。</param>
    internal WordEditors(MsWord.Editors editors)
    {
        _editors = editors ?? throw new ArgumentNullException(nameof(editors));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _editors != null ? new WordApplication(_editors.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _editors?.Parent;

    /// <inheritdoc/>
    public int Count => _editors?.Count ?? 0;

    /// <inheritdoc/>
    public IWordEditor this[object index]
    {
        get
        {
            if (_editors == null) return null;
            try
            {
                var comEditor = _editors.Item(index);
                return comEditor != null ? new WordEditor(comEditor) : null;
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
    public IWordEditor? Add(WdEditorType? editorID = null)
    {
        if (_editors == null) return null;
        try
        {
            var newEditor = _editors.Add(editorID.ComArgsConvert(e => e.EnumConvert(WdEditorType.wdEditorCurrent)));
            return newEditor != null ? new WordEditor(newEditor) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add editor: {ex.Message}", ex);
            return null;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordEditors"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _editors != null)
        {
            Marshal.ReleaseComObject(_editors);
            _editors = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordEditors"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordEditor> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordEditor> GetEnumerator()
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