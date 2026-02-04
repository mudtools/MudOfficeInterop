//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 对象中的标签集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint", NoneDisposed = false)]
public interface IPowerPointTags : IEnumerable<string?>, IDisposable
{
    /// <summary>
    /// 获取集合中的标签数量。
    /// </summary>
    /// <value>集合中的标签数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此标签集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此标签集合的父对象。
    /// </summary>
    /// <value>表示此标签集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过名称获取集合中的指定标签值。
    /// </summary>
    /// <param name="name">要获取的标签的名称。</param>
    /// <value>指定名称对应的标签值字符串。</value>
    string? this[string name] { get; }

    /// <summary>
    /// 在标签集合中添加新标签。
    /// </summary>
    /// <param name="name">新标签的名称。</param>
    /// <param name="value">新标签的值。</param>
    void Add(string name, string value);

    /// <summary>
    /// 删除指定名称的标签。
    /// </summary>
    /// <param name="name">要删除的标签的名称。</param>
    void Delete(string name);

    /// <summary>
    /// 添加二进制数据的标签。
    /// </summary>
    /// <param name="name">标签的名称。</param>
    /// <param name="filePath">包含二进制数据的文件路径。</param>
    void AddBinary(string name, string filePath);

    /// <summary>
    /// 获取二进制标签的值。
    /// </summary>
    /// <param name="name">二进制标签的名称。</param>
    /// <returns>二进制标签的值。</returns>
    int? BinaryValue(string name);

    /// <summary>
    /// 通过索引获取标签的名称。
    /// </summary>
    /// <param name="index">标签的索引（从1开始）。</param>
    /// <returns>指定索引处标签的名称。</returns>
    string? Name(int index);

    /// <summary>
    /// 通过索引获取标签的值。
    /// </summary>
    /// <param name="index">标签的索引（从1开始）。</param>
    /// <returns>指定索引处标签的值。</returns>
    string? Value(int index);
}