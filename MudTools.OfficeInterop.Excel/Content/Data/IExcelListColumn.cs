//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 表格（ListObject）中的一列，提供对列属性和操作的封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelListColumn : IDisposable
{
    /// <summary>
    /// 获取此列所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此列所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }


    IExcelListDataFormat ListDataFormat { get; }

    IExcelXPath XPath { get; }

    /// <summary>
    /// 获取此列在 ListColumns 集合中的索引（从 1 开始）。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置此列的标题名称（即表头）。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取此列对应的数据范围（DataBodyRange），不包含标题。
    /// 如果表无数据行，则可能为 null。
    /// </summary>
    IExcelRange? DataBodyRange { get; }

    /// <summary>
    /// 获取此列的整个范围（包括标题和数据体）。
    /// </summary>
    IExcelRange? Range { get; }

    /// <summary>
    /// 获取或设置此列的总计行公式（仅当 ListObject.ShowTotals = true 时有效）。
    /// </summary>
    XlTotalsCalculation TotalsCalculation { get; set; }

    /// <summary>
    /// 删除此列（将从表格中移除该列）。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此列对应的总计行单元格（仅当启用总计行时有效）。
    /// </summary>
    IExcelRange? Total { get; }
}
