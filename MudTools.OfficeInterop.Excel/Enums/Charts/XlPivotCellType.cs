//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 指定单元格对应的数据透视表实体类型。
/// </summary>
public enum XlPivotCellType
{
    /// <summary>
    /// 数据区域中的任意单元格（空白行除外）。
    /// </summary>
    xlPivotCellValue,

    /// <summary>
    /// 行或列区域中非小计、总计、自定义小计或空白行的单元格。
    /// </summary>
    xlPivotCellPivotItem,

    /// <summary>
    /// 行或列区域中的小计单元格。
    /// </summary>
    xlPivotCellSubtotal,

    /// <summary>
    /// 行或列区域中的总计单元格。
    /// </summary>
    xlPivotCellGrandTotal,

    /// <summary>
    /// 数据字段标签（非“数据”按钮）。
    /// </summary>
    xlPivotCellDataField,

    /// <summary>
    /// 字段按钮（非“数据”按钮）。
    /// </summary>
    xlPivotCellPivotField,

    /// <summary>
    /// 显示页字段所选项目的单元格。
    /// </summary>
    xlPivotCellPageFieldItem,

    /// <summary>
    /// 行或列区域中的自定义小计单元格。
    /// </summary>
    xlPivotCellCustomSubtotal,

    /// <summary>
    /// “数据”按钮。
    /// </summary>
    xlPivotCellDataPivotField,

    /// <summary>
    /// 数据透视表中的结构性空白单元格。
    /// </summary>
    xlPivotCellBlankCell
}
