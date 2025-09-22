//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SmartTagAction 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SmartTagAction 的安全访问和操作
/// </summary>
public interface IExcelSmartTagAction : IDisposable
{
    /// <summary>
    /// 获取智能标记动作的名称
    /// 对应 SmartTagAction.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取智能标记动作的名称
    /// 对应 SmartTagAction.Name 属性
    /// </summary>
    string TextboxText { get; }
    /// <summary>
    /// 获取智能标记动作的类型
    /// 对应 SmartTagAction.Type 属性
    /// </summary>
    XlSmartTagControlType Type { get; }

    /// <summary>
    /// 获取智能标记动作所在的智能标记对象
    /// 对应 SmartTagAction.Parent 属性
    /// </summary>
    IExcelSmartTag? Parent { get; }

    /// <summary>
    /// 执行该智能标记动作
    /// 对应 SmartTagAction.Execute 方法
    /// </summary>
    void Execute();
}