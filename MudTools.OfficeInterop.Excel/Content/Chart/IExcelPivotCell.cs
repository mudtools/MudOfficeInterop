//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 数据透视表中的一个单元格。该接口提供了访问数据透视表单元格各种属性和方法的功能，
/// 包括获取单元格的父对象、应用程序对象、单元格类型、关联的数据透视表、数据字段、透视表字段、透视表项等相关信息，以及处理单元格更改的方法。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotCell : IDisposable
{
    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据透视表单元格的类型。
    /// </summary>
    XlPivotCellType PivotCellType { get; }

    /// <summary>
    /// 获取包含此单元格的数据透视表。
    /// </summary>
    IExcelPivotTable? PivotTable { get; }

    /// <summary>
    /// 获取与此单元格关联的数据字段。
    /// </summary>
    IExcelPivotField? DataField { get; }

    /// <summary>
    /// 获取与此单元格关联的数据透视表字段。
    /// </summary>
    IExcelPivotField? PivotField { get; }

    /// <summary>
    /// 获取与此单元格关联的数据透视表项。
    /// </summary>
    IExcelPivotItem? PivotItem { get; }

    /// <summary>
    /// 获取与此单元格在同一行上的项目列表。
    /// </summary>
    IExcelPivotItemList? RowItems { get; }

    /// <summary>
    /// 获取与此单元格在同一列上的项目列表。
    /// </summary>
    IExcelPivotItemList? ColumnItems { get; }

    /// <summary>
    /// 获取与此数据透视表单元格对应的区域对象。
    /// </summary>
    IExcelRange? Range { get; }

    /// <summary>
    /// 获取自定义分类汇总函数。
    /// </summary>
    XlConsolidationFunction CustomSubtotalFunction { get; }

    /// <summary>
    /// 获取与此单元格关联的数据透视表行线。
    /// </summary>
    IExcelPivotLine? PivotRowLine { get; }

    /// <summary>
    /// 获取与此单元格关联的数据透视表列线。
    /// </summary>
    IExcelPivotLine? PivotColumnLine { get; }

    /// <summary>
    /// 分配更改到数据透视表单元格的值。
    /// </summary>
    void AllocateChange();

    /// <summary>
    /// 放弃对数据透视表单元格所做的更改。
    /// </summary>
    void DiscardChange();

    /// <summary>
    /// 获取数据源值。
    /// </summary>
    object DataSourceValue { get; }

    /// <summary>
    /// 获取单元格更改状态。
    /// </summary>
    XlCellChangedState CellChanged { get; }

    /// <summary>
    /// 获取多维表达式 (MDX) 名称。
    /// </summary>
    string MDX { get; }
}