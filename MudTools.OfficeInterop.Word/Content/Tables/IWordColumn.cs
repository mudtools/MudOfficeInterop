//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格列的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordColumn : IOfficeObject<IWordColumn>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取列索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置列宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取列单元格集合。
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 获取列边框。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取列底纹。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置列首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置列首选宽度类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 获取嵌套级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取一个值，该值指示此列是否为第一列。
    /// </summary>
    bool IsFirst { get; }

    /// <summary>
    /// 获取一个值，该值指示此列是否为最后一列。
    /// </summary>
    bool IsLast { get; }

    /// <summary>
    /// 获取下一个列对象。
    /// </summary>
    IWordColumn Next { get; }

    /// <summary>
    /// 获取上一个列对象。
    /// </summary>
    IWordColumn Previous { get; }

    /// <summary>
    /// 选择列。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除列。
    /// </summary>
    void Delete();

    /// <summary>
    /// 自动调整列宽以适应内容。
    /// </summary>
    void AutoFit();

    /// <summary>
    /// 设置列的宽度。
    /// </summary>
    /// <param name="ColumnWidth">要设置的列宽值</param>
    /// <param name="RulerStyle">标尺样式，用于确定如何解释宽度值</param>
    void SetWidth(float ColumnWidth, WdRulerStyle RulerStyle);

    /// <summary>
    /// 对列中的内容进行排序。
    /// </summary>
    /// <param name="excludeHeader">是否排除标题行进行排序</param>
    /// <param name="sortFieldType">排序字段类型</param>
    /// <param name="sortOrder">排序顺序（升序或降序等）</param>
    /// <param name="caseSensitive">是否区分大小写</param>
    /// <param name="bidiSort">是否使用双向排序</param>
    /// <param name="ignoreThe">是否忽略英文中的"The"</param>
    /// <param name="ignoreKashida">是否忽略阿拉伯语的kashida</param>
    /// <param name="ignoreDiacritics">是否忽略变音符号</param>
    /// <param name="ignoreHe">是否忽略希伯来语的"He"</param>
    /// <param name="languageID">用于排序的语言ID</param>
    void Sort(bool? excludeHeader = null, WdSortFieldType? sortFieldType = null, WdSortOrder? sortOrder = null,
              bool? caseSensitive = null, bool? bidiSort = null, bool? ignoreThe = null,
              bool? ignoreKashida = null, bool? ignoreDiacritics = null, bool? ignoreHe = null, WdLanguageID? languageID = null);

}