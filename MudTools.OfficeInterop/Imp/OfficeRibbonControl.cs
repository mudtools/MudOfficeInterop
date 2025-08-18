//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// Office RibbonControl 对象的二次封装实现类
/// 实现 IOfficeRibbonControl 接口
/// </summary>
internal class OfficeRibbonControl : IOfficeRibbonControl
{
    private MsCore.IRibbonControl _ribbonControl;
    // 如果需要与 IRibbonUI 交互以实现 Refresh，可能需要存储一个引用
    // private IOfficeRibbonUI _ribbonUI; // 或 Office.IRibbonUI _ribbonUI;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 OfficeRibbonControl 实例
    /// </summary>
    /// <param name="ribbonControl">要封装的 Microsoft.Office.Core.IRibbonControl 对象</param>
    /// <param name="ribbonUI">关联的 IRibbonUI 对象，用于刷新 (可选)</param>
    internal OfficeRibbonControl(MsCore.IRibbonControl ribbonControl /*, IOfficeRibbonUI ribbonUI = null*/)
    {
        _ribbonControl = ribbonControl ?? throw new ArgumentNullException(nameof(ribbonControl));
        // _ribbonUI = ribbonUI; // Store IRibbonUI if passed for Refresh functionality
    }

    #region 基础属性
    public string Id => _ribbonControl.Id;

    public object Context => _ribbonControl.Context;

    public object Tag => _ribbonControl.Tag;
    #endregion  

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _disposedValue = true;
        }
    }

    ~OfficeRibbonControl()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        // 由于 IRibbonControl 通常不应被释放，这里也应谨慎处理
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        // 不要更改此代码。将清理代码放入“Dispose(bool disposing)”方法中
        // 由于 IRibbonControl 通常不应被释放，这个方法主要处理包装类自身的资源
        Dispose(disposing: true);
        // SuppressFinalize is still called as good practice, even if the finalizer does little.
        GC.SuppressFinalize(this);
    }
    #endregion
}