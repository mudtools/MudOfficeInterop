//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Styles 的接口，用于操作文档样式集合。
/// </summary>
public interface IWordStyles : IEnumerable<IWordStyle>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取样式集合中的样式数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取样式（从1开始）。
    /// </summary>
    IWordStyle this[int index] { get; }

    /// <summary>
    /// 根据名称获取样式。
    /// </summary>
    IWordStyle this[string name] { get; }

    /// <summary>
    /// 添加新样式。
    /// </summary>
    /// <param name="name">样式名称。</param>
    /// <param name="type">样式类型。</param>
    /// <returns>新创建的样式。</returns>
    IWordStyle Add(string name, WdStyleType type = WdStyleType.wdStyleTypeParagraph);

    /// <summary>
    /// 检查样式是否存在。
    /// </summary>
    /// <param name="name">样式名称。</param>
    /// <returns>是否存在。</returns>
    bool Contains(string name);

    /// <summary>
    /// 根据样式类型获取样式列表。
    /// </summary>
    /// <param name="styleType">样式类型。</param>
    /// <returns>样式名称列表。</returns>
    List<string> GetStyleNamesByType(WdStyleType styleType);

    /// <summary>
    /// 获取所有样式名称。
    /// </summary>
    /// <returns>样式名称列表。</returns>
    List<string> GetAllStyleNames();

    /// <summary>
    /// 获取内置样式名称列表。
    /// </summary>
    /// <returns>内置样式名称列表。</returns>
    List<string> GetBuiltInStyleNames();

    /// <summary>
    /// 获取用户自定义样式名称列表。
    /// </summary>
    /// <returns>自定义样式名称列表。</returns>
    List<string> GetCustomStyleNames();

    /// <summary>
    /// 删除指定名称的样式。
    /// </summary>
    /// <param name="name">样式名称。</param>
    /// <returns>是否删除成功。</returns>
    bool DeleteStyle(string name);

    /// <summary>
    /// 清除所有用户自定义样式。
    /// </summary>
    void ClearCustomStyles();

    /// <summary>
    /// 获取默认段落样式。
    /// </summary>
    IWordStyle DefaultParagraphStyle { get; }

    /// <summary>
    /// 获取默认字符样式。
    /// </summary>
    IWordStyle DefaultCharacterStyle { get; }

}