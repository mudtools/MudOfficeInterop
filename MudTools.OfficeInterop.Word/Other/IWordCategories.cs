//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 中 Categories 集合的封装接口。
/// 该集合包含某一构建基块类型（如页眉、页脚）下的所有类别（Category）。
/// </summary>
public interface IWordCategories : IEnumerable<IWordCategory>, IDisposable
{
    /// <summary>
    /// 获取集合中类别的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 1-based 索引获取指定位置的类别。
    /// </summary>
    /// <param name="index">索引（从1开始）。</param>
    /// <returns>封装后的类别对象，若不存在则返回 null。</returns>
    IWordCategory? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的类别（区分大小写）。
    /// </summary>
    /// <param name="name">类别的名称。</param>
    /// <returns>封装后的类别对象，若不存在则返回 null。</returns>
    IWordCategory? this[string name] { get; }

    /// <summary>
    /// 判断是否存在指定名称的类别。
    /// </summary>
    /// <param name="name">要查找的类别名称。</param>
    /// <returns>存在返回 true，否则 false。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有类别的名称列表。
    /// </summary>
    /// <returns>类别名称字符串列表。</returns>
    List<string> GetNames();
}