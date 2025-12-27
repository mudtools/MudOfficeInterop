//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示一个内置对话框的封装实现类。
/// </summary>
internal class WordDialog : IWordDialog
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordDialog));
    private MsWord.Dialog _dialog;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordDialog"/> 类的新实例。
    /// </summary>
    /// <param name="dialog">要封装的原始 COM Dialog 对象。</param>
    internal WordDialog(MsWord.Dialog dialog)
    {
        _dialog = dialog ?? throw new ArgumentNullException(nameof(dialog));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication? Application => _dialog != null ? new WordApplication(_dialog.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _dialog?.Parent;

    /// <inheritdoc/>
    public int Creator => _dialog?.Creator ?? 0;

    #endregion

    #region 对话框属性实现 (Dialog Properties Implementation)

    /// <inheritdoc/>
    public int CommandBarId => _dialog?.CommandBarId ?? 0;

    /// <inheritdoc/>
    public string CommandName => _dialog?.CommandName ?? string.Empty;

    /// <inheritdoc/>
    public WdWordDialogTab DefaultTab
    {
        get => _dialog?.DefaultTab != null ? (WdWordDialogTab)(int)_dialog?.DefaultTab : WdWordDialogTab.wdDialogOrganizerTabAutoText;
        set
        {
            if (_dialog != null) _dialog.DefaultTab = (MsWord.WdWordDialogTab)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdWordDialog Type => _dialog?.Type != null ? (WdWordDialog)(int)_dialog?.Type : WdWordDialog.wdDialogHelpAbout;

    #endregion

    #region 对话框方法实现 (Dialog Methods Implementation)

    /// <inheritdoc/>
    public bool Display(float? timeout = null)
    {
        if (_dialog == null) return false;
        try
        {
            // 注意：Display 方法返回 -1 (True) 或 0 (False)，需要转换
            var result = _dialog.Display(timeout.ComArgsVal());
            return result != 0; // Word 返回 -1 表示 True
        }
        catch (COMException ex)
        {
            log.Error($"Failed to display dialog: {ex.Message}");
            return false;
        }
    }

    /// <inheritdoc/>
    public void Execute()
    {
        _dialog?.Execute();
    }

    /// <inheritdoc/>
    public bool Show(float? timeout = null)
    {
        if (_dialog == null) return false;
        try
        {
            // 注意：Show 方法返回 -1 (True) 或 0 (False)，需要转换
            var result = _dialog.Show(timeout.ComArgsVal());
            return result != 0; // Word 返回 -1 表示 True
        }
        catch (COMException ex)
        {
            log.Error($"Failed to show dialog: {ex.Message}");
            return false;
        }
    }

    /// <inheritdoc/>
    public void Update()
    {
        _dialog?.Update();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordDialog"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dialog != null)
        {
            Marshal.ReleaseComObject(_dialog);
            _dialog = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordDialog"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}