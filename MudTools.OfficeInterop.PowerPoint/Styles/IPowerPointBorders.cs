//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// 表示一组边框的集合，通常用于表格单元格或形状。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointBorders : IOfficeObject<IPowerPointBorders, MsPowerPoint.Borders>, IEnumerable<IPowerPointLineFormat?>, IDisposable
{
    /// <summary>
    /// 获取集合中边框的数量。
    /// </summary>
    /// <value>集合中边框的数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此边框集合的 PowerPoint 应用程序对象。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此边框集合的父对象。
    /// </summary>
    /// <value>表示父对象的 <see cref="object"/>。</value>
    object Parent { get; }

    /// <summary>
    /// 根据边框类型获取对应的边框线条格式。
    /// </summary>
    /// <param name="borderType">要获取的边框类型。</param>
    /// <value>指定边框类型的 <see cref="IPowerPointLineFormat"/> 对象。</value>
    IPowerPointLineFormat? this[PpBorderType borderType] { get; }
}