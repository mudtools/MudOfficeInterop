//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示容器文档的一个自定义或内置文档属性。
/// 此接口是对 Microsoft.Office.Core.DocumentProperty COM 对象的二次封装。
/// </summary>
public interface IOfficeDocumentProperty : IDisposable
{
    /// <summary>
    /// 获取文档属性的应用程序对象。
    /// </summary>
    object Application { get; }

    /// <summary>
    /// 获取一个 32 位整数，用于指示创建该对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置文档属性的名称。
    /// 对于内置属性，此属性为只读；对于自定义属性，此属性可读写。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取文档属性的类型。
    /// 对于内置属性，此属性为只读；对于自定义属性，此属性可读写。
    /// 类型由 MsoDocProperties 枚举定义（如字符串、数字、日期、布尔值等）[[25]]。
    /// </summary>
    MsoDocProperties Type { get; set; }

    /// <summary>
    /// 获取或设置文档属性的值。
    /// 值的类型必须与 <see cref="Type"/> 属性匹配。
    /// </summary>
    object Value { get; set; }

    /// <summary>
    /// 获取一个值，该值指示文档属性是否为内置属性。
    /// </summary>
    bool IsBuiltIn { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示自定义文档属性的值是否链接到容器文档的内容。
    /// 此属性仅对自定义文档属性有效。如果为 true，则属性值会随链接内容动态更新 [[44]]。
    /// </summary>
    bool LinkToContent { get; set; }

    /// <summary>
    /// 删除自定义文档属性。
    /// 此方法不能用于删除内置文档属性。
    /// </summary>
    void Delete();
}
