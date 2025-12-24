//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 文档中的脚本对象接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeScript : IDisposable
{
    /// <summary>
    /// 获取脚本对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置脚本的扩展属性
    /// </summary>
    string Extended { get; set; }

    /// <summary>
    /// 获取或设置脚本的唯一标识符
    /// </summary>
    string Id { get; set; }

    /// <summary>
    /// 获取或设置脚本的文本内容
    /// </summary>
    string ScriptText { get; set; }

    /// <summary>
    /// 获取或设置脚本的语言类型
    /// </summary>
    MsoScriptLanguage Language { get; set; }

    /// <summary>
    /// 获取脚本在文档中的位置
    /// </summary>
    MsoScriptLocation Location { get; }

    /// <summary>
    /// 删除当前脚本对象
    /// </summary>
    void Delete();
}