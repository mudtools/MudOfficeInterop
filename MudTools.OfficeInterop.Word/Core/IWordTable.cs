//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中表格的封装接口。
/// </summary>
public interface IWordTable : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表格所在的范围。
    /// </summary>
    IWordRange? Range { get; }

    bool Uniform { get; }

    /// <summary>
    /// 获取表格的行集合。
    /// </summary>
    IWordRows? Rows { get; }

    /// <summary>
    /// 获取表格的列集合。
    /// </summary>
    IWordColumns? Columns { get; }

    /// <summary>
    /// 获取表格的边框集合。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取嵌套在当前表格中的表格集合。
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取表格的底纹设置。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置表格是否允许跨页断行。
    /// </summary>
    bool AllowPageBreaks { get; set; }

    /// <summary>
    /// 获取或设置表格是否允许调整大小。
    /// </summary>
    bool AllowAutoFit { get; set; }

    /// <summary>
    /// 获取或设置表格的样式。
    /// </summary>
    object? TableStyle { get; set; }

    /// <summary>
    /// 获取或设置表格标题行。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置表格的描述文字。
    /// </summary>
    string Descr { get; set; }

    /// <summary>
    /// 获取表格的嵌套层数。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的标题行应用样式。
    /// </summary>
    bool ApplyStyleHeadingRows { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的最后一行应用样式。
    /// </summary>
    bool ApplyStyleLastRow { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的第一列应用样式。
    /// </summary>
    bool ApplyStyleFirstColumn { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的最后一列应用样式。
    /// </summary>
    bool ApplyStyleLastColumn { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的行带应用样式。
    /// </summary>
    bool ApplyStyleRowBands { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对表格的列带应用样式。
    /// </summary>
    bool ApplyStyleColumnBands { get; set; }

    /// <summary>
    /// 获取或设置表格的首选宽度值。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置表格首选宽度的类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 获取或设置表格中文字的排列方向。
    /// </summary>
    WdTableDirection TableDirection { get; set; }

    /// <summary>
    /// 获取或设置表格中单元格之间的间距。
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置表格单元格右边距。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取或设置表格单元格左边距。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取或设置表格单元格下边距。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取或设置表格单元格上边距。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取或设置表格的ID。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 通过行列索引获取单元格。
    /// </summary>
    /// <param name="rowIndex">行索引（从1开始）。</param>
    /// <param name="columnIndex">列索引（从1开始）。</param>
    IWordCell? Cell(int rowIndex, int columnIndex);


    /// <summary>
    /// 转换表格为文本。
    /// </summary>
    /// <param name="separator">分隔符。</param>
    /// <returns>转换后的范围。</returns>
    IWordRange? ConvertToText(object separator);

    /// <summary>
    /// 删除表格。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选中表格。
    /// </summary>
    void Select();

    /// <summary>
    /// 拆分表格。
    /// </summary>
    /// <param name="beforeRow">在指定行之前拆分。</param>
    void Split(object beforeRow);

    /// <summary>
    /// 排序表格。
    /// </summary>
    /// <param name="excludeHeader">是否排除标题行。</param>
    /// <param name="fieldNumber">排序字段号。</param>
    /// <param name="sortFieldType">排序字段类型。</param>
    /// <param name="ascending">是否升序。</param>
    void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object ascending);

    /// <summary>
    /// 应用表格样式。
    /// </summary>
    /// <param name="styleName">样式名称。</param>
    void ApplyStyleDirectFormatting(string styleName);
}