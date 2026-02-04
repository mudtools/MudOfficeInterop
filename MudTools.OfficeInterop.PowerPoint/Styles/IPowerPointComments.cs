//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片中的注释集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointComments : IOfficeObject<IPowerPointComments, MsPowerPoint.Comments>, IEnumerable<IPowerPointComment?>, IDisposable
{

    /// <summary>
    /// 获取集合中的注释数量。
    /// </summary>
    /// <value>集合中的注释数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此注释集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此注释集合的父对象。
    /// </summary>
    /// <value>表示此注释集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引获取集合中的指定注释。
    /// </summary>
    /// <param name="index">要获取的注释的索引（从1开始）。</param>
    /// <value>位于指定索引处的 <see cref="IPowerPointComment"/> 对象。</value>
    IPowerPointComment? this[int index] { get; }

    /// <summary>
    /// 在注释集合中添加新注释。
    /// </summary>
    /// <param name="left">注释在幻灯片上的左边缘位置（以磅为单位）。</param>
    /// <param name="top">注释在幻灯片上的上边缘位置（以磅为单位）。</param>
    /// <param name="author">注释的作者姓名。</param>
    /// <param name="authorInitials">注释的作者首字母缩写。</param>
    /// <param name="text">注释的文本内容。</param>
    /// <returns>新添加的 <see cref="IPowerPointComment"/> 对象。</returns>
    IPowerPointComment? Add(float left, float top, string author, string authorInitials, string text);

    /// <summary>
    /// 在注释集合中添加新注释（增强版本）。
    /// </summary>
    /// <param name="left">注释在幻灯片上的左边缘位置（以磅为单位）。</param>
    /// <param name="top">注释在幻灯片上的上边缘位置（以磅为单位）。</param>
    /// <param name="author">注释的作者姓名。</param>
    /// <param name="authorInitials">注释的作者首字母缩写。</param>
    /// <param name="text">注释的文本内容。</param>
    /// <param name="providerID">注释提供程序的标识符。</param>
    /// <param name="userID">注释用户的标识符。</param>
    /// <returns>新添加的 <see cref="IPowerPointComment"/> 对象。</returns>
    IPowerPointComment? Add2(float left, float top, string author, string authorInitials, string text, string providerID, string userID);
}