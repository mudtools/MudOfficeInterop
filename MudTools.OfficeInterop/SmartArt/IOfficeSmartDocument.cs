//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Microsoft Office Word 2003 Document 对象和 Microsoft Office Excel 2003 Workbook 对象的 SmartDocument 属性返回一个 SmartDocument 对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeSmartDocument : IOfficeObject<IOfficeSmartDocument>, IDisposable
{
    /// <summary>
    /// 获取或设置标识附加到活动 Microsoft Office Word 2003 文档或 Microsoft Office Excel 2003 工作簿的 XML 扩展包的 ID（通常是全局唯一标识符 GUID）。
    /// </summary>
    string SolutionID { get; set; }

    /// <summary>
    /// 获取或设置提供附加到活动 Microsoft Office Word 2003 文档或 Microsoft Office Excel 2003 工作簿的 XML 扩展包文件的完整路径的绝对 URL。
    /// </summary>
    string SolutionURL { get; set; }

    /// <summary>
    /// 显示一个对话框，允许用户选择可用的 XML 扩展包以附加到活动 Microsoft Office Word 2003 文档或 Microsoft Office Excel 2003 工作簿。
    /// </summary>
    /// <param name="considerAllSchemas">可选 Boolean。True 显示用户计算机上安装的所有可用 XML 扩展包。False 仅显示适用于活动文档的 XML 扩展包。默认值为 False。</param>
    void PickSolution(bool considerAllSchemas = false);

    /// <summary>
    /// 刷新活动 Microsoft Office Word 2003 文档或 Microsoft Office Excel 2003 工作簿的文档操作任务窗格。
    /// </summary>
    void RefreshPane();

}