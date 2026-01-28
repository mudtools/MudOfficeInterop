//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 表格中单元格的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCellRange : IEnumerable<IPowerPointCell>, IDisposable
{
    /// <summary>
    /// 获取集合中的单元格数量。
    /// </summary>
    /// <value>集合中的项数。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此单元格范围的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此单元格范围的父对象。
    /// </summary>
    /// <value>表示此单元格范围父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引获取集合中的指定单元格。
    /// </summary>
    /// <param name="index">要获取的单元格的索引（从1开始）。</param>
    /// <value>位于指定索引处的 <see cref="IPowerPointCell"/> 对象。</value>
    IPowerPointCell? this[int index] { get; }

    /// <summary>
    /// 获取单元格范围的所有边框。
    /// </summary>
    /// <value>表示单元格边框集合的 <see cref="IPowerPointBorders"/> 对象。</value>
    IPowerPointBorders? Borders { get; }
}