//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ErrorCheckingOptions 对象的二次封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelErrorCheckingOptions : IDisposable
{
    /// <summary>
    /// 返回指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回表示 Microsoft Excel 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否为所有违反启用的错误检查规则的单元格向用户发出警报。
    /// </summary>
    bool BackgroundChecking { get; set; }

    /// <summary>
    /// 获取或设置错误检查选项指示器的颜色。
    /// </summary>
    XlColorIndex IndicatorColorIndex { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 用自动更正选项按钮标识包含计算结果为错误的公式的选定单元格。False 禁用对计算结果为错误值的单元格的错误检查。
    /// </summary>
    bool EvaluateToError { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 用自动更正选项按钮标识包含带两位年份的文本日期的单元格。False 禁用对包含带两位年份的文本日期的单元格的错误检查。
    /// </summary>
    bool TextDate { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 用自动更正选项按钮标识包含写作文本的数字的选定单元格。False 禁用对写作文本的数字的错误检查。
    /// </summary>
    bool NumberAsText { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 标识区域中包含不一致公式的单元格。False 禁用不一致公式检查。
    /// </summary>
    bool InconsistentFormula { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 用自动更正选项按钮标识引用省略了可包括的相邻单元格的范围的公式的选定单元格。False 禁用对省略单元格的错误检查。
    /// </summary>
    bool OmittedCells { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 标识已解锁且包含公式的选定单元格。False 禁用对包含公式的解锁单元格的错误检查。
    /// </summary>
    bool UnlockedFormulaCells { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当设置为 True（默认）时，Microsoft Excel 用自动更正选项按钮标识引用空单元格的公式的选定单元格。False 禁用空单元格引用检查。
    /// </summary>
    bool EmptyCellReferences { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，如果列表中的数据验证已启用，则为 True。
    /// </summary>
    bool ListDataValidation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示表格公式是否不一致。
    /// </summary>
    bool InconsistentTableFormula { get; set; }
}