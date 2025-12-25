//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft.Office.Core.CustomXMLValidationErrors 集合中的单个验证错误。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLValidationError : IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLValidationError 对象的父对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object Parent { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLValidationError 对象中错误的名称。如果不存在错误，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>String</returns>
    string Name { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLValidationError 对象中的节点（如果存在）。如果不存在节点，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode Node { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLValidationError 对象中的文本。只读。
    /// </summary>
    /// <returns>String</returns>
    string Text { get; }

    /// <summary>
    /// 获取从 Microsoft.Office.Core.CustomXMLValidationError 对象生成的错误类型。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.MsoCustomXMLValidationErrorType</returns>
    MsoCustomXMLValidationErrorType Type { get; }

    /// <summary>
    /// 获取表示 Microsoft.Office.Core.CustomXMLValidationError 对象中验证错误的数字。只读。
    /// </summary>
    /// <returns>Integer</returns>
    int ErrorCode { get; }

    /// <summary>
    /// 删除表示数据验证错误的 Microsoft.Office.Core.CustomXMLValidationError 对象。
    /// </summary>
    void Delete();
}