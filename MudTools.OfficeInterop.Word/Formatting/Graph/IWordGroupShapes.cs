//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示组合形状中的单个形状的集合。
/// <para>注：每个形状由 Shape 对象表示。使用 Item[Object] 方法可在组中处理单个形状，而无需取消组合它们。</para>
/// <para>使用 GroupItems 属性可返回 GroupShapes 集合。</para>
/// </summary>
public interface IWordGroupShapes : IEnumerable<IWordShape>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的形状数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的单个形状对象。
    /// </summary>
    /// <param name="index">形状的索引（从 1 开始）或名称。</param>
    /// <returns>指定的形状对象。</returns>
    IWordShape this[object index] { get; }

    /// <summary>
    /// 返回一个 ShapeRange 对象，该对象代表 GroupShapes 集合中的指定子集。
    /// </summary>
    /// <param name="index">Variant 类型，指定要包含在范围内的单个对象或包含对象的数组。</param>
    /// <returns>指定的形状范围。</returns>
    IWordShapeRange Range(object index);
}