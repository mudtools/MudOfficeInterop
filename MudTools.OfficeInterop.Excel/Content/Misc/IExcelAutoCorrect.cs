//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 包含 Microsoft Excel 自动更正属性（日期名称的大写、纠正两个首字母大写、自动更正列表等）。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelAutoCorrect : IOfficeObject<IExcelAutoCorrect, MsExcel.AutoCorrect>, IDisposable
{
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 向自动更正替换数组中添加一个条目。
    /// </summary>
    /// <param name="what">要被替换的文本。如果此字符串已存在于自动更正替换数组中，现有的替换文本将被新文本替换。</param>
    /// <param name="replacement">替换文本。</param>
    object AddReplacement(string what, string replacement);

    /// <summary>
    /// 获取或设置一个值，指示是否自动将日期名称的首字母大写。
    /// </summary>
    bool CapitalizeNamesOfDays { get; set; }

    /// <summary>
    /// 从自动更正替换数组中删除一个条目。
    /// </summary>
    /// <param name="what">要被替换的文本，如自动更正替换数组中要删除的行所示。如果此字符串不存在于自动更正替换数组中，则此方法失败。</param>
    object DeleteReplacement(string what);

    /// <summary>
    /// 获取或设置自动更正替换数组。
    /// </summary>
    /// <param name="index">要返回的自动更正替换数组的行索引。行以包含两个元素的一维数组形式返回：第一个元素是第 1 列中的文本，第二个元素是第 2 列中的文本。</param>
    [MethodIndex]
    object ReplacementList(int index);

    /// <summary>
    /// 获取或设置自动更正替换数组。
    /// </summary>
    [ComPropertyWrap(PropertyName = "ReplacementList")]
    object ReplacementLists { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否自动替换自动更正替换列表中的文本。
    /// </summary>
    bool ReplaceText { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否自动纠正以两个大写字母开头的单词。
    /// </summary>
    bool TwoInitialCapitals { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否自动纠正句子（第一个单词）的大写。
    /// </summary>
    bool CorrectSentenceCap { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否自动纠正意外使用 CAPS LOCK 键。
    /// </summary>
    bool CorrectCapsLock { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示自动更正选项按钮。默认值为 true。
    /// </summary>
    bool DisplayAutoCorrectOptions { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为列表启用自动扩展。
    /// </summary>
    bool AutoExpandListRange { get; set; }

    /// <summary>
    /// 获取或设置一个值，影响通过自动填充列表创建的计算列的创建。
    /// </summary>
    bool AutoFillFormulasInLists { get; set; }
}