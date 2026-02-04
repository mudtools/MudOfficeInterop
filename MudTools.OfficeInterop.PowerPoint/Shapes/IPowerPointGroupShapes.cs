//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示组合形状中的子形状集合，提供对组合内各个形状的访问。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointGroupShapes : IOfficeObject<IPowerPointGroupShapes, MsPowerPoint.GroupShapes>, IEnumerable<IPowerPointShape?>, IDisposable
{
    /// <summary>
    /// 获取创建此组合形状集合的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取组合形状集合的父对象。
    /// </summary>
    /// <value>父对象，通常是组合形状。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取组合中形状的数量。
    /// </summary>
    /// <value>组合中子形状的总数。</value>
    int Count { get; }

    /// <summary>
    /// 通过索引获取组合中的形状。
    /// </summary>
    /// <param name="index">要获取的形状的索引或名称。</param>
    /// <returns>指定索引处的形状。</returns>
    IPowerPointShape? this[int index] { get; }

    /// <summary>
    /// 通过索引获取组合中的形状。
    /// </summary>
    /// <param name="index">要获取的形状的索引或名称。</param>
    /// <returns>指定索引处的形状。</returns>
    IPowerPointShape? this[string index] { get; }

    /// <summary>
    /// 获取一个形状范围，包含组合中指定的形状。
    /// </summary>
    /// <param name="index">形状的索引、名称或索引数组。</param>
    /// <returns>包含指定形状的形状范围对象。</returns>
    IPowerPointShapeRange? Range(object index);
}