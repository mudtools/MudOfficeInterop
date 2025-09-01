//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office RibbonUI 对象的二次封装实现类
/// 实现 IOfficeRibbonUI 接口
/// </summary>
internal class OfficeRibbonUI : IOfficeRibbonUI
{
    private MsCore.IRibbonUI _ribbonUI;
    private object _application;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 OfficeRibbonUI 实例
    /// </summary>
    /// <param name="ribbonUI">要封装的 Microsoft.Office.Core.IRibbonUI 对象</param>
    /// <param name="application">关联的 Application 对象 (可选)</param>
    internal OfficeRibbonUI(MsCore.IRibbonUI ribbonUI, object application = null)
    {
        _ribbonUI = ribbonUI ?? throw new ArgumentNullException(nameof(ribbonUI));
        _application = application;
    }

    #region 基础属性

    public object Application
    {
        get
        {
            return _application;
        }
    }
    #endregion

    #region 操作方法
    public void Invalidate()
    {
        _ribbonUI.Invalidate();
    }

    public void InvalidateControl(string controlID)
    {
        _ribbonUI.InvalidateControl(controlID);
    }
    #endregion

    #region 高级功能 
    public void ActivateTab(string tabId)
    {
        _ribbonUI.ActivateTab(tabId);
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (_ribbonUI != null)
            {
                try
                {
                    while (Marshal.ReleaseComObject(_ribbonUI) > 0) { }
                }
                catch
                {
                }
                _ribbonUI = null;
            }

            _disposedValue = true;
        }
    }

    ~OfficeRibbonUI()
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