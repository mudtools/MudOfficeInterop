//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档或节中文本列的集合。
/// 封装了 Microsoft.Office.Interop.Word.TextColumns 对象。
/// </summary>
/// <remarks>
/// 使用 PageSetup 对象的 TextColumns 属性可返回 TextColumns 集合。
/// </remarks>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordTextColumns : IEnumerable<IWordTextColumn?>, IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordTextColumns"/> 对象的父对象。
    /// </summary>
    /// <remarks>父对象通常是 PageSetup 对象。</remarks>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中文本列的计数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置文本列之间所需的间距（以磅为单位）。
    /// </summary>
    /// <remarks>
    /// 设置此属性会将所有文本列设置为等宽，并在各列之间创建指定的间距。
    /// 在使用此属性之前，必须将 LineBetween 属性设置为 false。
    /// </remarks>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，该值指示是否在文本列之间添加垂直线。
    /// </summary>
    int LineBetween { get; set; }

    /// <summary>
    /// 获取或设置文本列的内容流动方向。
    /// </summary>
    /// <value>WdFlowDirection 枚举值，指定文本列的流动方向。</value>
    WdFlowDirection FlowDirection { get; set; }

    /// <summary>
    /// 获取或设置文本列的宽度（以磅为单位）。
    /// </summary>
    /// <value>float 类型，表示文本列的宽度。</value>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示文本列是否等间距分布。
    /// </summary>
    /// <value>int 类型，非零值表示列将等间距分布，零值表示列将保留各自的宽度。</value>
    int EvenlySpaced { get; set; }

    /// <summary>
    /// 通过索引（从 1 开始）获取集合中的单个文本列。
    /// </summary>
    /// <param name="index">文本列的索引号。</param>
    /// <returns>指定索引处的 <see cref="IWordTextColumn"/> 对象。</returns>
    /// <exception cref="ArgumentOutOfRangeException">如果索引超出范围。</exception>
    IWordTextColumn? this[int index] { get; }

    #endregion // 属性

    #region 方法

    /// <summary>
    /// 向文本列集合中添加一列。
    /// </summary>
    /// <param name="width">
    /// 文本列的宽度（以磅为单位）。
    /// 如果省略此参数，则添加一列，其余列宽进行调整以适应页面宽度。
    /// </param>
    /// <param name="spacing">
    /// 文本列与前一列之间的间距（以磅为单位）。
    /// 如果省略此参数，则使用 Spacing 属性的值。
    /// </param>
    /// <param name="evenlySpaced">
    /// 如果为 True，则使各文本列等宽，同时保留指定的列间距或 Spacing 属性的值。
    /// 如果为 False，则保留原来各列的宽度，并在前一列的右边添加新列。
    /// 如果省略此参数，则假定其值为 True。
    /// </param>
    /// <returns>新添加的 <see cref="IWordTextColumn"/> 对象。</returns>
    IWordTextColumn? Add(int? width, int? spacing, int? evenlySpaced);

    /// <summary>
    /// 设置指定文档或节的文本列数。
    /// </summary>
    /// <param name="count">所需的文本列数。</param>
    /// <remarks>
    /// 此方法会将现有列替换为指定数量的新列。
    /// 新列具有相等的宽度和间距。
    /// </remarks>
    void SetCount(int count);

    #endregion
}