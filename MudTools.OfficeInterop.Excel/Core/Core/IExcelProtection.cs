//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表的各种保护选项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelProtection : IOfficeObject<IExcelProtection, MsExcel.Protection>, IDisposable
{
    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许设置单元格格式。
    /// </summary>
    bool AllowFormattingCells { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许设置列格式。
    /// </summary>
    bool AllowFormattingColumns { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许设置行格式。
    /// </summary>
    bool AllowFormattingRows { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许插入列。
    /// </summary>
    bool AllowInsertingColumns { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许插入行。
    /// </summary>
    bool AllowInsertingRows { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许插入超链接。
    /// </summary>
    bool AllowInsertingHyperlinks { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许删除列。
    /// </summary>
    bool AllowDeletingColumns { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许删除行。
    /// </summary>
    bool AllowDeletingRows { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许排序。
    /// </summary>
    bool AllowSorting { get; }

    /// <summary>
    /// 获取一个值，该值指示在保护工作表之前创建的自动筛选是否允许用户使用。
    /// </summary>
    bool AllowFiltering { get; }

    /// <summary>
    /// 获取一个值，该值指示在受保护的工作表上是否允许用户操作数据透视表。
    /// </summary>
    bool AllowUsingPivotTables { get; }

    /// <summary>
    /// 获取 AllowEditRanges 对象，该对象表示工作表中允许编辑的区域。
    /// </summary>
    IExcelAllowEditRanges? AllowEditRanges { get; }
}