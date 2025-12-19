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
internal partial class OfficeCustomTaskPane
{
    private MsCore._CustomTaskPaneEvents_Event _customTaskPaneEvents_Event;

    public event EventHandler<TaskPaneVisibleStateChangedEventArgs> VisibleStateChanged;
    public event EventHandler<TaskPaneDockPositionChangedEventArgs> DockPositionChanged;

    /// <summary>
    /// 初始化 OfficeCustomTaskPane 实例
    /// </summary>
    /// <param name="customTaskPane">要封装的 Microsoft.Office.Core.CustomTaskPane 对象</param>
    internal OfficeCustomTaskPane(MsCore.CustomTaskPane customTaskPane)
    {
        _customTaskPane = customTaskPane ?? throw new ArgumentNullException(nameof(customTaskPane));
        _customTaskPaneEvents_Event = customTaskPane;
        ConectEvent();
    }

    private void ConectEvent()
    {
        if (_customTaskPaneEvents_Event != null)
        {
            _customTaskPaneEvents_Event.VisibleStateChange += OnComVisibleStateChanged;
            _customTaskPaneEvents_Event.DockPositionStateChange += OnComDockPositionChanged;
        }
    }

    private void DisConnectEvent()
    {
        if (_customTaskPaneEvents_Event != null)
        {
            _customTaskPaneEvents_Event.VisibleStateChange -= OnComVisibleStateChanged;
            _customTaskPaneEvents_Event.DockPositionStateChange -= OnComDockPositionChanged;
        }
    }

    private void OnComVisibleStateChanged(MsCore.CustomTaskPane customTaskPane)
    {
        VisibleStateChanged?.Invoke(this, new TaskPaneVisibleStateChangedEventArgs(customTaskPane.Visible));
    }

    private void OnComDockPositionChanged(MsCore.CustomTaskPane customTaskPane)
    {
        DockPositionChanged?.Invoke(this, new TaskPaneDockPositionChangedEventArgs((MsoCTPDockPosition)customTaskPane.DockPosition));
    }

    #region IDisposable Support
    protected void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing && _customTaskPane != null)
        {
            DisConnectEvent();
            Marshal.ReleaseComObject(_customTaskPane);
            _customTaskPane = null;
            _customTaskPaneEvents_Event = null;
        }
        _disposedValue = true;
    }
    #endregion
}
