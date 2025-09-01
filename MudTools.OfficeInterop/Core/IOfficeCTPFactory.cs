//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office CTP (Custom Task Pane) Factory 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.ICTPFactory 的安全访问和操作
/// ICTPFactory 对象用于创建 CustomTaskPane 对象
/// </summary>
public interface IOfficeCTPFactory : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 CTP 工厂所在的Application对象
    /// 注意：标准 ICTPFactory 接口没有直接的 Application 属性。
    /// 此处使用 object 作为通用占位符，或在实现中通过上下文关联。
    /// </summary>
    object Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 创建一个新的自定义任务窗格 (CustomTaskPane)
    /// 对应 ICTPFactory.CreateCTP 方法 (或 GetInstance, depending on exact interface)
    /// </summary>
    /// <param name="CTPAxID ProgID">用于任务窗格内容的 COM 对象的 ProgID (例如，Windows Forms UserControl 的 ProgID)</param>
    /// <param name="title">任务窗格的标题</param>
    /// <param name="CTPParentWindow">任务窗格的父容器</param>
    /// <returns>新创建的自定义任务窗格对象</returns>
    IOfficeCustomTaskPane CreateCTP(string CTPAxID, string title, object CTPParentWindow);
    #endregion

}
