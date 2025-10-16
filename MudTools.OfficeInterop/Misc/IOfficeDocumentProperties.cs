//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop;


/// <summary>
/// 表示 DocumentProperty 对象的集合。
/// 此接口是对 Microsoft.Office.Core.DocumentProperties COM 对象的二次封装。
/// </summary>
public interface IOfficeDocumentProperties : IEnumerable<IOfficeDocumentProperty>, IDisposable
{
    /// <summary>
    /// 获取集合中的文档属性总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的文档属性。
    /// 索引从 1 开始。
    /// </summary>
    /// <param name="index">文档属性的索引（从1开始）。</param>
    /// <returns>指定索引处的文档属性，如果未找到则返回 null。</returns>
    IOfficeDocumentProperty? this[int index] { get; }

    /// <summary>
    /// 获取集合中具有指定名称的文档属性。
    /// </summary>
    /// <param name="name">文档属性的名称。</param>
    /// <returns>具有指定名称的文档属性，如果未找到则返回 null。</returns>
    IOfficeDocumentProperty? this[string name] { get; }

    /// <summary>
    /// 向集合中添加一个新的自定义文档属性。
    /// </summary>
    /// <param name="name">新自定义属性的名称。</param>
    /// <param name="linkToContent">如果为 true，则新属性将链接到文档内容；否则为静态值。</param>
    /// <param name="type">新属性的数据类型（MsoDocProperties 枚举）。</param>
    /// <param name="value">新属性的值。如果 <paramref name="linkToContent"/> 为 true，则此参数应为链接源（如Excel中的单元格地址）。</param>
    /// <param name="linkSource">（可选）当 <paramref name="linkToContent"/> 为 true 时，指定链接源。对于Excel，这通常是单元格地址或区域名称。</param>
    /// <returns>新创建的文档属性，如果创建失败则返回 null。</returns>
    IOfficeDocumentProperty? Add(string name, bool linkToContent, MsoDocProperties type, object value, object? linkSource = null);
}