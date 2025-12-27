//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 列表格式的封装接口。
/// </summary>
public interface IWordListFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置列表级别数。
    /// </summary>
    int ListLevelNumber { get; set; }

    IWordListTemplate? ListTemplate { get; }

    /// <summary>
    /// 获取或设置列表类型。
    /// </summary>
    WdListType ListType { get; }

    /// <summary>
    /// 获取列表字符串。
    /// </summary>
    string ListString { get; }

    /// <summary>
    /// 获取或设置是否为单个列表。
    /// </summary>
    bool SingleList { get; }

    /// <summary>
    /// 获取或设置是否为单个列表模板。
    /// </summary>
    bool SingleListTemplate { get; }

    /// <summary>
    /// 应用默认的项目符号格式。
    /// </summary>
    /// <param name="DefaultListBehavior">指定列表行为的默认设置。</param>
    void ApplyBulletDefault(WdDefaultListBehavior? DefaultListBehavior = null);

    /// <summary>
    /// 应用默认的数字编号格式。
    /// </summary>
    /// <param name="DefaultListBehavior">指定列表行为的默认设置。</param>
    void ApplyNumberDefault(WdDefaultListBehavior? DefaultListBehavior = null);

    /// <summary>
    /// 应用默认的大纲编号格式。
    /// </summary>
    /// <param name="DefaultListBehavior">指定列表行为的默认设置。</param>
    void ApplyOutlineNumberDefault(WdDefaultListBehavior? DefaultListBehavior = null);

    /// <summary>
    /// 应用列表模板。
    /// </summary>
    /// <param name="listTemplate">列表模板。</param>
    /// <param name="continuePreviousList">是否继续前一个列表。</param>
    /// <param name="applyTo">应用到的范围。</param>
    /// <param name="defaultListBehavior">默认列表行为。</param>
    void ApplyListTemplateWithLevel(IWordListTemplate listTemplate, bool continuePreviousList,
                                  WdListApplyTo applyTo, WdDefaultListBehavior defaultListBehavior);


    /// <summary>
    /// 应用列表模板。
    /// </summary>
    /// <param name="listTemplate">列表模板。</param>
    /// <param name="continuePreviousList">是否继续前一个列表。</param>
    /// <param name="applyTo">应用到的范围。</param>
    /// <param name="defaultListBehavior">默认列表行为。</param>
    void ApplyListTemplate(IWordListTemplate listTemplate, bool continuePreviousList,
                                  WdListApplyTo applyTo, WdDefaultListBehavior defaultListBehavior);

    /// <summary>
    /// 移除列表格式。
    /// </summary>
    void RemoveNumbers();

    /// <summary>
    /// 检查是否可以列表还原。
    /// </summary>
    /// <returns>是否可以还原。</returns>
    WdContinue CanContinuePreviousList(IWordListTemplate listTemplate);

    /// <summary>
    /// 列表降级。
    /// </summary>
    void ListIndent();

    /// <summary>
    /// 列表升级。
    /// </summary>
    void ListOutdent();
}