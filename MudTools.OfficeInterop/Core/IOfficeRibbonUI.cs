//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office RibbonUI 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.IRibbonUI 的安全访问和操作
/// IRibbonUI 对象通常在自定义功能区加载时通过回调函数传递
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore", ComClassName = "IRibbonUI")]
public interface IOfficeRibbonUI : IDisposable
{
    #region 操作方法
    /// <summary>
    /// 使整个功能区用户界面无效并强制重新绘制
    /// 对应 IRibbonUI.Invalidate 方法
    /// 调用此方法后，所有 GetEnabled、GetVisible、GetLabel 等回调函数将被重新调用
    /// </summary>
    void Invalidate();

    /// <summary>
    /// 使功能区上指定标识符的控件无效并强制重新绘制
    /// 对应 IRibbonUI.InvalidateControl 方法
    /// 调用此方法后，指定控件的 GetEnabled、GetVisible、GetLabel 等回调函数将被重新调用
    /// </summary>
    /// <param name="controlID">要无效的控件的标识符 (id)</param>
    void InvalidateControl(string controlID);

    /// <summary>
    /// 使功能区上指定标识符的内置控件无效并强制重新绘制
    /// 对应 IRibbonUI.InvalidateControlMSO 方法
    /// 调用此方法后，指定控件的 GetEnabled、GetVisible、GetLabel 等回调函数将被重新调用
    /// </summary>
    /// <param name="controlID">要无效的内置控件的标识符</param>
    void InvalidateControlMso(string controlID);

    /// <summary>
    /// 激活功能区上的某个特定选项卡
    /// 注意：标准 IRibbonUI 接口没有此方法。这通常通过设置 ActiveTab 属性在 customUI 标签中实现，
    /// 或者通过 Application.COMAddIns 或其他方式间接实现。
    /// 此处作为高级功能占位符。
    /// </summary>
    /// <param name="tabId">要激活的选项卡的标识符</param>
    void ActivateTab(string tabId);

    /// <summary>
    /// 激活功能区上具有指定命名空间的特定选项卡
    /// 这是 ActivateTab 方法的增强版本，支持命名空间限定的控件标识符
    /// </summary>
    /// <param name="controlID">要激活的选项卡的标识符</param>
    /// <param name="nameSpace">选项卡控件所属的命名空间</param>
    void ActivateTabQ(string controlID, string nameSpace);
    #endregion
}