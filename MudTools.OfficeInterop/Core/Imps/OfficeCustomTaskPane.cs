//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office CustomTaskPane 对象的二次封装实现类
/// 实现 IOfficeCustomTaskPane 接口
/// </summary>
internal class OfficeCustomTaskPane : IOfficeCustomTaskPane
{
    private MsCore.CustomTaskPane _customTaskPane;
    private bool _disposedValue = false;

    public event EventHandler<TaskPaneVisibleStateChangedEventArgs> VisibleStateChanged;
    public event EventHandler<TaskPaneDockPositionChangedEventArgs> DockPositionChanged;

    /// <summary>
    /// 初始化 OfficeCustomTaskPane 实例
    /// </summary>
    /// <param name="customTaskPane">要封装的 Microsoft.Office.Core.CustomTaskPane 对象</param>
    internal OfficeCustomTaskPane(MsCore.CustomTaskPane customTaskPane)
    {
        _customTaskPane = customTaskPane ?? throw new ArgumentNullException(nameof(customTaskPane));
        if (_customTaskPane != null)
        {
            _customTaskPane.VisibleStateChange += OnComVisibleStateChanged;
            _customTaskPane.DockPositionStateChange += OnComDockPositionChanged;
        }
    }

    private void OnComVisibleStateChanged(MsCore.CustomTaskPane customTaskPane)
    {
        VisibleStateChanged?.Invoke(this, new TaskPaneVisibleStateChangedEventArgs(customTaskPane.Visible));
    }

    private void OnComDockPositionChanged(MsCore.CustomTaskPane customTaskPane)
    {
        DockPositionChanged?.Invoke(this, new TaskPaneDockPositionChangedEventArgs((MsoDockPosition)customTaskPane.DockPosition));
    }

    #region 基础属性
    public string Title => _customTaskPane.Title;

    public object Application => _customTaskPane.Application; // 通用占位符

    public object Window => _customTaskPane.Window; // 通用占位符

    public bool Visible
    {
        get => _customTaskPane.Visible;
        set => _customTaskPane.Visible = value;
    }

    public object ContentControl => _customTaskPane.ContentControl;

    public MsoDockPosition DockPosition
    {
        get => (MsoDockPosition)_customTaskPane.DockPosition;
        set => _customTaskPane.DockPosition = (MsCore.MsoCTPDockPosition)value;
    }

    public MsoDockPositionRestrict DockPositionRestrict
    {
        get => (MsoDockPositionRestrict)_customTaskPane.DockPositionRestrict;
        set => _customTaskPane.DockPositionRestrict = (MsCore.MsoCTPDockPositionRestrict)value;
    }
    #endregion

    #region 位置和大小
    public int Width
    {
        get => _customTaskPane.Width;
        set => _customTaskPane.Width = value;
    }

    public int Height
    {
        get => _customTaskPane.Height;
        set => _customTaskPane.Height = value;
    }

    #endregion

    #region 操作方法
    public void Delete()
    {
        _customTaskPane.Delete();
    }
    #endregion

    #region 高级功能 (概念性)
    public void Refresh()
    {
        bool currentVis = this.Visible;
        this.Visible = !currentVis;
        this.Visible = currentVis;
    }

    public void ActivateContent()
    {

    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (_customTaskPane != null)
        {
            _customTaskPane.VisibleStateChange -= OnComVisibleStateChanged;
            _customTaskPane.DockPositionStateChange -= OnComDockPositionChanged;

            Marshal.ReleaseComObject(_customTaskPane);
            _customTaskPane = null;
        }
        _disposedValue = true;
    }

    ~OfficeCustomTaskPane()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
