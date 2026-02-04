//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 演示文稿中的幻灯片集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlides : IEnumerable<IPowerPointSlide?>, IDisposable
{
    /// <summary>
    /// 获取集合中的幻灯片数量。
    /// </summary>
    /// <value>集合中的幻灯片数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此幻灯片集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片集合的父对象。
    /// </summary>
    /// <value>表示此幻灯片集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定幻灯片。
    /// </summary>
    /// <param name="index">要获取的幻灯片的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointSlide"/> 对象。</value>
    IPowerPointSlide? this[int index] { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定幻灯片。
    /// </summary>
    /// <param name="name">要获取的幻灯片的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointSlide"/> 对象。</value>
    IPowerPointSlide? this[string name] { get; }

    /// <summary>
    /// 通过幻灯片标识符查找幻灯片。
    /// </summary>
    /// <param name="slideID">要查找的幻灯片的标识符。</param>
    /// <returns>找到的 <see cref="IPowerPointSlide"/> 对象。</returns>
    IPowerPointSlide? FindBySlideID(int slideID);

    /// <summary>
    /// 在幻灯片集合中添加新幻灯片。
    /// </summary>
    /// <param name="index">新幻灯片要插入的位置索引（从1开始）。</param>
    /// <param name="layout">新幻灯片的版式。</param>
    /// <returns>新添加的 <see cref="IPowerPointSlide"/> 对象。</returns>
    IPowerPointSlide? Add(int index, PpSlideLayout layout);

    /// <summary>
    /// 从文件插入幻灯片到集合中。
    /// </summary>
    /// <param name="fileName">包含幻灯片的文件名称。</param>
    /// <param name="index">插入幻灯片的位置索引。</param>
    /// <param name="slideStart">要插入的起始幻灯片编号。</param>
    /// <param name="slideEnd">要插入的结束幻灯片编号。值为-1表示插入所有幻灯片。</param>
    /// <returns>实际插入的幻灯片数量。</returns>
    int? InsertFromFile(string fileName, int index, int slideStart = 1, int slideEnd = -1);

    /// <summary>
    /// 获取指定幻灯片的范围。
    /// </summary>
    /// <param name="index">要获取范围的幻灯片的索引（从1开始）、索引数组或名称。</param>
    /// <returns>指定幻灯片范围对应的 <see cref="IPowerPointSlideRange"/> 对象。</returns>
    IPowerPointSlideRange? Range(object index);

    /// <summary>
    /// 粘贴剪贴板内容为幻灯片。
    /// </summary>
    /// <param name="index">粘贴幻灯片的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>粘贴的幻灯片范围。</returns>
    IPowerPointSlideRange? Paste(int index = -1);

    /// <summary>
    /// 使用自定义版式添加新幻灯片。
    /// </summary>
    /// <param name="index">新幻灯片要插入的位置索引（从1开始）。</param>
    /// <param name="pCustomLayout">要使用的自定义版式。</param>
    /// <returns>新添加的 <see cref="IPowerPointSlide"/> 对象。</returns>
    IPowerPointSlide? AddSlide(int index, IPowerPointCustomLayout pCustomLayout);
}