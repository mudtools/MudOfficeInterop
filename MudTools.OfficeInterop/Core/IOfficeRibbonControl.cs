//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office RibbonControl 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.IRibbonControl 的安全访问
/// IRibbonControl 对象通常在自定义功能区控件的回调函数中作为参数传递
/// </summary>
public interface IOfficeRibbonControl : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取功能区控件的标识符 (id)
    /// 对应 IRibbonControl.Id 属性
    /// </summary>
    string Id { get; }

    /// <summary>
    /// 获取与功能区控件关联的上下文对象
    /// 对应 IRibbonControl.Context 属性
    /// 这通常是触发控件操作的对象，例如 Window 对象、Range 对象等
    /// </summary>
    object Context { get; }

    /// <summary>
    /// 获取与功能区控件关联的标签对象
    /// 对应 IRibbonControl.Tag 属性
    /// 这是在 XML 中定义控件时指定的任意数据
    /// </summary>
    object Tag { get; }
    #endregion
}