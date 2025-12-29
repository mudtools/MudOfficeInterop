//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示文档中的一个内容控件。
/// <para>内容控件是文档中绑定的、可能添加标签的区域，它们充当特定类型内容（如日期、列表或格式化文本段落）的容器。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordContentControl : IOfficeObject<IWordContentControl>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置内容控件的标题。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置内容控件的标签。
    /// </summary>
    string Tag { get; set; }

    /// <summary>
    /// 获取或设置内容控件的类型。
    /// </summary>
    WdContentControlType Type { get; }

    /// <summary>
    /// 获取内容控件的范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置内容控件的默认文本。
    /// </summary>
    IWordBuildingBlock? PlaceholderText { get; }

    /// <summary>
    /// 获取或设置内容控件是否被锁定，禁止用户编辑其内容。
    /// </summary>
    bool LockContentControl { get; set; }

    /// <summary>
    /// 获取或设置内容控件的内容是否被锁定，禁止用户编辑其内容（但可以删除整个控件）。
    /// </summary>
    bool LockContents { get; set; }

    /// <summary>
    /// 获取或设置内容控件的临时状态。临时内容控件在用户编辑其内容时不会被保存。
    /// </summary>
    bool Temporary { get; set; }

    /// <summary>
    /// 获取或设置内容控件的 XML 映射。
    /// </summary>
    IWordXMLMapping? XMLMapping { get; }

    /// <summary>
    /// 获取内容控件的下拉列表条目（适用于 ComboBox 和 DropdownList 类型）。
    /// </summary>
    IWordContentControlListEntries? DropdownListEntries { get; }

    /// <summary>
    /// 获取内容控件的日期显示格式（适用于 Date 类型）。
    /// </summary>
    string DateDisplayFormat { get; set; }

    /// <summary>
    /// 获取内容控件的日期存储格式（适用于 Date 类型）。
    /// </summary>
    WdContentControlDateStorageFormat DateStorageFormat { get; set; }

    /// <summary>
    /// 获取或设置内容控件的日期显示语言（适用于 Date 类型）。
    /// </summary>
    WdLanguageID DateDisplayLocale { get; set; }

    /// <summary>
    /// 获取或设置内容控件的多行状态（适用于 RichText 和 PlainText 类型）。
    /// </summary>
    bool MultiLine { get; set; }

    /// <summary>
    /// 获取内容控件的 ID。
    /// </summary>
    string ID { get; }

    /// <summary>
    /// 获取或设置内容控件的复选框状态（适用于 CheckBox 类型）。
    /// </summary>
    bool Checked { get; set; }

    /// <summary>
    /// 获取内容控件的父内容控件（如果存在）。
    /// </summary>
    IWordContentControl? ParentContentControl { get; }

    /// <summary>
    /// 删除此内容控件及其内容。
    /// </summary>
    /// <param name="deleteContents">如果为 true，则同时删除控件内容；如果为 false，则仅删除控件本身。</param>
    void Delete(bool deleteContents);

    /// <summary>
    /// 将内容控件复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 设置内容控件的已选中符号。
    /// </summary>
    void SetUncheckedSymbol(int characterNumber, string font = "");

    /// <summary>
    /// 设置内容控件的未选中符号。
    /// </summary>
    /// <param name="characterNumber"></param>
    /// <param name="font"></param>
    void SetCheckedSymbol(int characterNumber, string font = "");
}
