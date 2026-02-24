//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 应用程序中所有打开的演示文稿集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPresentations : IOfficeObject<IPowerPointPresentations, MsPowerPoint.Presentations>, IEnumerable<IPowerPointPresentation?>, IDisposable
{

    /// <summary>
    /// 获取集合中的演示文稿数量。
    /// </summary>
    /// <value>集合中的演示文稿数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此演示文稿集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此演示文稿集合的父对象。
    /// </summary>
    /// <value>表示此演示文稿集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定演示文稿。
    /// </summary>
    /// <param name="index">要获取的演示文稿的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointPresentation"/> 对象。</value>
    IPowerPointPresentation? this[int index] { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定演示文稿。
    /// </summary>
    /// <param name="name">要获取的演示文稿的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointPresentation"/> 对象。</value>
    IPowerPointPresentation? this[string name] { get; }

    /// <summary>
    /// 创建新的演示文稿。
    /// </summary>
    /// <param name="withWindow">指示是否在新窗口中显示的布尔值。</param>
    /// <returns>新创建的 <see cref="IPowerPointPresentation"/> 对象。</returns>
    IPowerPointPresentation? Add([ConvertTriState] bool withWindow = true);

    /// <summary>
    /// 打开指定的演示文稿文件。
    /// </summary>
    /// <param name="fileName">要打开的演示文稿文件名。</param>
    /// <param name="readOnly">指示是否以只读方式打开的布尔值。</param>
    /// <param name="untitled">指示是否作为无标题文档打开的布尔值。</param>
    /// <param name="withWindow">指示是否在新窗口中显示的布尔值。</param>
    /// <returns>打开的 <see cref="IPowerPointPresentation"/> 对象。</returns>
    IPowerPointPresentation? Open(string fileName, [ConvertTriState] bool readOnly = false, [ConvertTriState] bool untitled = false, [ConvertTriState] bool withWindow = true);

    /// <summary>
    /// 检出指定的演示文稿文件。
    /// </summary>
    /// <param name="fileName">要检出的演示文稿文件名。</param>
    void CheckOut(string fileName);

    /// <summary>
    /// 检查是否可以检出指定的演示文稿文件。
    /// </summary>
    /// <param name="fileName">要检查的演示文稿文件名。</param>
    /// <returns>如果可以检出，则返回 true；否则返回 false。</returns>
    bool? CanCheckOut(string fileName);

    /// <summary>
    /// 打开指定的演示文稿文件（支持修复功能）。
    /// </summary>
    /// <param name="fileName">要打开的演示文稿文件名。</param>
    /// <param name="readOnly">指示是否以只读方式打开的布尔值。</param>
    /// <param name="untitled">指示是否作为无标题文档打开的布尔值。</param>
    /// <param name="withWindow">指示是否在新窗口中显示的布尔值。</param>
    /// <param name="openAndRepair">指示是否打开并尝试修复的布尔值。</param>
    /// <returns>打开的 <see cref="IPowerPointPresentation"/> 对象。</returns>
    IPowerPointPresentation? Open2007(string fileName, [ConvertTriState] bool readOnly = false, [ConvertTriState] bool untitled = false, [ConvertTriState] bool withWindow = true, [ConvertTriState] bool openAndRepair = false);
}