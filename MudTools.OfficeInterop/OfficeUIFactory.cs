//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop;


/// <summary>
/// Office用户界面工厂类，用于创建各种Office UI对象的包装器实例
/// 提供对CTP工厂、功能区UI和功能区控件等Office核心UI组件的统一创建接口
/// </summary>
public static class OfficeUIFactory
{

    public static T? Create<T>(this T t, object comObj) where T : IOfficeObject<T>
    {
        return t.LoadFromObject(comObj);
    }

    /// <summary>
    /// 创建自定义任务窗格工厂的包装器实例
    /// </summary>
    /// <param name="officeCTPFactory">原始的Office CTP工厂对象</param>
    /// <returns>IOfficeCTPFactory接口实例，如果传入的参数为null则返回null</returns>
    public static IOfficeCTPFactory? CreateCTPFactory(MsCore.ICTPFactory officeCTPFactory)
    {
        if (officeCTPFactory == null) return null;
        return new OfficeCTPFactory(officeCTPFactory);
    }

    /// <summary>
    /// 创建功能区UI的包装器实例
    /// </summary>
    /// <param name="ribbonUI">原始的Office功能区UI对象</param>
    /// <returns>IOfficeRibbonUI接口实例，如果传入的参数为null则返回null</returns>
    public static IOfficeRibbonUI? CreateRibbonUI(MsCore.IRibbonUI ribbonUI)
    {
        if (ribbonUI == null)
            return null;
        return new OfficeRibbonUI(ribbonUI);
    }

    /// <summary>
    /// 创建功能区控件的包装器实例
    /// </summary>
    /// <param name="ribbonControl">原始的Office功能区控件对象</param>
    /// <returns>IOfficeRibbonControl接口实例，如果传入的参数为null则返回null</returns>
    public static IOfficeRibbonControl CreateRibbonControl(MsCore.IRibbonControl ribbonControl)
    {
        if (ribbonControl == null) return null;
        return new OfficeRibbonControl(ribbonControl);
    }
}