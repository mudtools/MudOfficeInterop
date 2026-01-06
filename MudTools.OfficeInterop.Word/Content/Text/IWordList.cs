//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示已应用于文档中指定段落的单个列表格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordList : IOfficeObject<IWordList, MsWord.List>, IDisposable
{

    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取表示指定对象中包含的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取表示列表中所有编号段落的 ListParagraphs 集合。
    /// </summary>
    IWordListParagraphs? ListParagraphs { get; }

    /// <summary>
    /// 获取一个值，指示整个列表对象是否使用相同的列表模板。
    /// </summary>
    bool SingleListTemplate { get; }

    /// <summary>
    /// 获取指示创建指定对象的应用程序的 32 位整数。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 将指定列表中的列表编号和 LISTNUM 字段转换为文本。
    /// </summary>
    /// <param name="numberType">可选项。要转换的数字类型。可以是任何 WdNumberType 常量。</param>
    void ConvertNumbersToText(WdNumberType? numberType = null);

    /// <summary>
    /// 从指定列表中移除编号或项目符号。
    /// </summary>
    /// <param name="numberType">可选项。WdNumberType。要移除的数字类型。</param>
    void RemoveNumbers(WdNumberType? numberType = null);

    /// <summary>
    /// 返回指定列表中项目符号或编号项目以及 LISTNUM 字段的数量。
    /// </summary>
    /// <param name="numberType">可选项。要计数的数字类型。可以是以下 WdNumberType 常量之一：wdNumberParagraph、wdNumberListNum 或 wdNumberAllNumbers。默认值为 wdNumberAllNumbers。</param>
    /// <param name="level">可选项。对应于要计数的编号级别的数字。如果省略此参数，则计数所有级别。</param>
    /// <returns>指定列表中的项目数量。</returns>
    int? CountNumberedItems(WdNumberType? numberType = null, int? level = null);

    /// <summary>
    /// 返回一个 WdContinue 常量（wdContinueDisabled、wdResetList 或 wdContinueList），指示是否可以继续上一个列表的格式。
    /// </summary>
    /// <param name="listTemplate">必需。ListTemplate 对象。已应用于文档中先前段落的列表模板。</param>
    /// <returns>表示是否可以继续上一个列表格式的常量。</returns>
    WdContinue? CanContinuePreviousList(IWordListTemplate listTemplate);

    /// <summary>
    /// 将一组列表格式特性应用于指定列表。
    /// </summary>
    /// <param name="listTemplate">必需。ListTemplate 对象。要应用的列表模板。</param>
    /// <param name="continuePreviousList">可选项。True 表示继续上一个列表的编号；False 表示开始新列表。</param>
    /// <param name="defaultListBehavior">可选项。设置一个值，指定 Microsoft Word 是否使用新的面向 Web 的格式以获得更好的列表显示。可以是以下常量之一：wdWord8ListBehavior（使用与 Microsoft Word 97 兼容的格式）或 wdWord9ListBehavior（使用面向 Web 的格式）。为了兼容性，默认常量为 wdWord8ListBehavior，但在新过程中应使用 wdWord9ListBehavior，以利用缩进和多级列表方面改进的面向 Web 的格式。</param>
    void ApplyListTemplate(IWordListTemplate listTemplate, bool? continuePreviousList = null, WdDefaultListBehavior? defaultListBehavior = null);

    /// <summary>
    /// 获取应用于指定自动图文集条目的样式名称。
    /// </summary>
    string StyleName { get; }

    /// <summary>
    /// 将一组列表格式特性应用于指定列表，可选地指定应用级别。
    /// </summary>
    /// <param name="listTemplate">要应用的列表模板。</param>
    /// <param name="continuePreviousList">True 表示继续上一个列表的编号；False 表示开始新列表。</param>
    /// <param name="defaultListBehavior">设置一个值，指定 Microsoft Word 是否使用新的面向 Web 的格式以获得更好的列表显示。可以是以下 WdDefaultListBehavior 值之一：WdDefaultListBehavior.wdWord8ListBehavior（使用与 Microsoft Word 97 兼容的格式）或 WdDefaultListBehavior.wdWord9ListBehavior（使用面向 Web 的格式）。为了兼容性，默认常量为 WdDefaultListBehavior.wdWord8ListBehavior，但在新过程中应使用 WdDefaultListBehavior.wdWord9ListBehavior，以利用缩进和多级列表方面改进的面向 Web 的格式。</param>
    /// <param name="applyLevel">要应用列表模板的级别。</param>
    void ApplyListTemplateWithLevel(IWordListTemplate listTemplate, bool? continuePreviousList = null, WdDefaultListBehavior? defaultListBehavior = null, int? applyLevel = null);


}