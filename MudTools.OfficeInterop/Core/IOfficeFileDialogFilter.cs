//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// Office FileDialogFilter 对象的二次封装接口
/// 提供对 Microsoft.Office.Core.FileDialogFilter 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeFileDialogFilter : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取文件对话框过滤器对象的父对象 (通常是 FileDialogFilters 集合)
    /// 对应 FileDialogFilter.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取过滤器的描述文本 (例如 "Text Files")
    /// 对应 FileDialogFilter.Description 属性
    /// </summary>
    string Description { get; }

    /// <summary>
    /// 获取过滤器的扩展名模式 (例如 "*.txt" 或 "*.txt;*.csv")
    /// 对应 FileDialogFilter.Extensions 属性
    /// </summary>
    string Extensions { get; }
    #endregion
}
