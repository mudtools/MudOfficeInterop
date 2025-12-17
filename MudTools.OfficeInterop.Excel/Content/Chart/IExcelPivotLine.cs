//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 中的数据透视表线条对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotLine : IDisposable
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
    /// 获取数据透视表线条的类型。
    /// </summary>
    XlPivotLineType LineType { get; }

    /// <summary>
    /// 获取数据透视表线条的位置。
    /// </summary>
    int Position { get; }

    /// <summary>
    /// 获取数据透视表线条单元格集合。
    /// </summary>
    IExcelPivotLineCells PivotLineCells { get; }

    /// <summary>
    /// 获取完整的数据透视表线条单元格集合。
    /// </summary>
    IExcelPivotLineCells PivotLineCellsFull { get; }
}