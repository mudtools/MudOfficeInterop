//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordProtectedViewWindow : IWordProtectedViewWindow
{
    private MsWord.ProtectedViewWindow _protectedViewWindow;
    private bool _disposedValue;

    internal WordProtectedViewWindow(MsWord.ProtectedViewWindow protectedViewWindow)
    {
        _protectedViewWindow = protectedViewWindow ?? throw new ArgumentNullException(nameof(protectedViewWindow));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _protectedViewWindow?.Application != null ? new WordApplication(_protectedViewWindow.Application) : null;

    /// <summary>
    /// 获取代表指定对象的父对象的对象。
    /// </summary>
    public object Parent => _protectedViewWindow?.Parent;

    /// <summary>
    /// 获取受保护的视图窗口的窗口标题。
    /// </summary>
    public string Caption => _protectedViewWindow?.Caption;

    /// <summary>
    /// 获取受保护的视图窗口的高度（以磅为单位）。
    /// </summary>
    public int Height
    {
        get => _protectedViewWindow?.Height ?? 0;
        set
        {
            if (_protectedViewWindow != null)
                _protectedViewWindow.Height = value;
        }
    }

    /// <summary>
    /// 获取或设置受保护的视图窗口的水平位置（以磅为单位）。
    /// </summary>
    public int Left
    {
        get => _protectedViewWindow?.Left ?? 0;
        set
        {
            if (_protectedViewWindow != null)
                _protectedViewWindow.Left = value;
        }
    }

    /// <summary>
    /// 获取受保护的视图窗口的垂直位置（以磅为单位）。
    /// </summary>
    public int Top
    {
        get => _protectedViewWindow?.Top ?? 0;
        set
        {
            if (_protectedViewWindow != null)
                _protectedViewWindow.Top = value;
        }
    }
    /// <summary>
    /// 获取或设置受保护的视图窗口的宽度（以磅为单位）。
    /// </summary>
    public int Width
    {
        get => _protectedViewWindow?.Width ?? 0;
        set
        {
            if (_protectedViewWindow != null)
                _protectedViewWindow.Width = value;
        }
    }

    /// <summary>
    /// 获取或设置受保护的视图窗口的窗口状态。
    /// </summary>
    public WdWindowState WindowState
    {
        get => _protectedViewWindow?.WindowState != null ? (WdWindowState)(int)_protectedViewWindow?.WindowState : WdWindowState.wdWindowStateNormal;
        set
        {
            if (_protectedViewWindow != null) _protectedViewWindow.WindowState = (MsWord.WdWindowState)(int)value;
        }
    }

    /// <summary>
    /// 获取受保护的视图窗口的显示状态。
    /// </summary>
    public bool Visible
    {
        get => _protectedViewWindow?.Visible ?? false;
        set
        {
            if (_protectedViewWindow != null)
                _protectedViewWindow.Visible = value;
        }
    }

    /// <summary>
    /// 获取受保护的视图窗口所显示的文档的完整路径。
    /// </summary>
    public string DocumentFullName => _protectedViewWindow?.Document?.FullName;

    /// <summary>
    /// 获取受保护的视图窗口所显示的文档对象。
    /// </summary>
    public IWordDocument? Document =>
       _protectedViewWindow?.Document != null ? new WordDocument(_protectedViewWindow.Document) : null;

    /// <summary>
    /// 获取受保护的视图窗口的唯一标识符。
    /// </summary>
    public int Index => _protectedViewWindow?.Index != null ? _protectedViewWindow.Index : -1;

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值代表是否激活受保护的视图窗口。
    /// </summary>
    public bool Active
    {
        get => _protectedViewWindow?.Active ?? false;
    }

    public string SourceName
    {
        get => _protectedViewWindow?.SourceName ?? string.Empty;
    }

    public string SourcePath
    {
        get => _protectedViewWindow?.SourcePath ?? string.Empty;
    }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    public int Creator => _protectedViewWindow?.Creator ?? 0;

    #endregion

    #region 方法实现

    /// <summary>
    /// 关闭受保护的视图窗口。
    /// </summary>
    public void Close()
    {
        _protectedViewWindow?.Close();
    }

    /// <summary>
    /// 编辑受保护的视图窗口中的文档。
    /// </summary>
    /// <param name="PasswordTemplate">打开模板时所需的密码。</param>
    /// <param name="WritePasswordDocument">打开文档时所需的写入密码。</param>
    /// <param name="WritePasswordTemplate">打开模板时所需的写入密码。</param>
    /// <returns>返回编辑中的文档对象。</returns>
    public IWordDocument? Edit(string PasswordTemplate, string WritePasswordDocument, string WritePasswordTemplate)
    {
        var doc = _protectedViewWindow?.Edit(PasswordTemplate, WritePasswordDocument, WritePasswordTemplate);
        return doc != null ? new WordDocument(doc) : null;
    }

    /// <summary>
    /// 激活受保护的视图窗口。
    /// </summary>
    public void Activate()
    {
        _protectedViewWindow?.Activate();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _protectedViewWindow != null)
        {
            Marshal.ReleaseComObject(_protectedViewWindow);
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