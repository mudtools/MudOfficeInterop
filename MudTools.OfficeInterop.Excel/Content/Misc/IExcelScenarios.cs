//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示指定工作表上所有 Scenario 对象的集合。方案是一组已命名并保存的输入值（称为可变单元格）。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelScenarios : IEnumerable<IExcelScenario?>, IOfficeObject<IExcelScenarios, MsExcel.Scenarios>, IDisposable
{
    /// <summary>
    /// 获取对象所在的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取最近使用的文件集合中的文件数量
    /// 对应 RecentFiles.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// 索引从1开始
    /// </summary>
    /// <param name="index">对象索引（从1开始）</param>
    /// <returns>最近使用文件对象</returns>
    IExcelScenario? this[int index] { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="index">必需。对象的名称或索引号。</param>
    /// <returns>指定索引处的 Scenario 对象。</returns>
    IExcelScenario? this[string index] { get; }

    /// <summary>
    /// 创建一个新工作表，其中包含指定工作表上方案的摘要报告。
    /// </summary>
    /// <param name="reportType">可选项。XlSummaryReportType。报告类型。</param>
    /// <param name="resultCells">可选项。Range 对象，表示指定工作表上的结果单元格。通常，此区域引用一个或多个包含依赖于模型中可变单元格值的公式的单元格，即显示特定方案结果的单元格。如果省略此参数，报告中不包含结果单元格。</param>
    void CreateSummary(XlSummaryReportType reportType = XlSummaryReportType.xlStandardSummary, IExcelRange? resultCells = null);

    /// <summary>
    /// 将另一个工作表中的方案合并到当前 Scenarios 集合中。
    /// </summary>
    /// <param name="source">必需。包含要合并的方案的工作表的名称，或表示该工作表的 Worksheet 对象。</param>
    void Merge(string source);

    /// <summary>
    /// 将另一个工作表中的方案合并到当前 Scenarios 集合中。
    /// </summary>
    /// <param name="source">必需。包含要合并的方案的工作表的名称，或表示该工作表的 Worksheet 对象。</param>
    void Merge(IExcelWorksheet source);
}