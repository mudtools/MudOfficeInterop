//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中一组边框（Borders）的封装接口。
/// <para>注：Borders 集合的成员数量是有限的，并且取决于父对象的类型。</para>
/// </summary>
public interface IWordBorders : IEnumerable<IWordBorder>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取集合中的边框数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（<see cref="MsWord.WdBorderType"/> 常量）获取单个边框。
    /// </summary>
    /// <param name="index">标识边框的 <see cref="MsWord.WdBorderType"/> 常量。</param>
    /// <returns>指定的边框对象。</returns>
    IWordBorder? this[WdBorderType index] { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否启用所有边框的格式。
    /// </summary>
    bool Enable { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否删除段落和表格边缘的垂直边框，以便水平边框可以连接到页面边框。
    /// </summary>
    bool JoinBorders { get; set; }

    /// <summary>
    /// 获取或设置内部边框的 24 位颜色。
    /// </summary>
    WdColor InsideColor { get; set; }

    /// <summary>
    /// 获取或设置内部边框的颜色索引。
    /// </summary>
    WdColorIndex InsideColorIndex { get; set; }

    /// <summary>
    /// 获取或设置内部边框的线条样式。
    /// </summary>
    WdLineStyle InsideLineStyle { get; set; }

    /// <summary>
    /// 获取或设置内部边框的线条宽度。
    /// </summary>
    WdLineWidth InsideLineWidth { get; set; }

    /// <summary>
    /// 获取或设置外部边框的 24 位颜色。
    /// </summary>
    WdColor OutsideColor { get; set; }

    /// <summary>
    /// 获取或设置外部边框的颜色索引。
    /// </summary>
    WdColorIndex OutsideColorIndex { get; set; }

    /// <summary>
    /// 获取或设置外部边框的线条样式。
    /// </summary>
    WdLineStyle OutsideLineStyle { get; set; }

    /// <summary>
    /// 获取或设置外部边框的线条宽度。
    /// </summary>
    WdLineWidth OutsideLineWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否可以将水平边框应用于对象。
    /// </summary>
    bool HasHorizontal { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否可以将垂直边框应用于对象。
    /// </summary>
    bool HasVertical { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示页面边框是否显示在文档文本的前面。
    /// </summary>
    bool AlwaysInFront { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页面边框从页面边缘测量还是从环绕的文本测量。
    /// </summary>
    WdBorderDistanceFrom DistanceFrom { get; set; }

    /// <summary>
    /// 获取或设置文本与下边框之间的间距（以磅为单位）。
    /// </summary>
    int DistanceFromBottom { get; set; }

    /// <summary>
    /// 获取或设置文本与左边框之间的间距（以磅为单位）。
    /// </summary>
    int DistanceFromLeft { get; set; }

    /// <summary>
    /// 获取或设置文本与右边框之间的间距（以磅为单位）。
    /// </summary>
    int DistanceFromRight { get; set; }

    /// <summary>
    /// 获取或设置文本与上边框之间的间距（以磅为单位）。
    /// </summary>
    int DistanceFromTop { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否为节中的第一页启用了页面边框。
    /// </summary>
    bool EnableFirstPageInSection { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否为节中的所有页面（第一页除外）启用了页面边框。
    /// </summary>
    bool EnableOtherPagesInSection { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页面边框是否包含文档页脚。
    /// </summary>
    bool SurroundFooter { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页面边框是否包含文档页眉。
    /// </summary>
    bool SurroundHeader { get; set; }

    /// <summary>
    /// 将指定的页面边框格式应用于文档中的所有节。
    /// </summary>
    void ApplyPageBordersToAllSections();
}