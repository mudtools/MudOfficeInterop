//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Excel;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PivotCache 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotCache 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotCache : IOfficeObject<IExcelPivotCache, MsExcel.PivotCache>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表缓存的索引位置
    /// 对应 PivotCache.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取数据透视表缓存的父对象 (通常是 Workbook)
    /// 对应 PivotCache.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据透视表缓存所在的Application对象
    /// 对应 PivotCache.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据透视表缓存的源数据类型
    /// 对应 PivotCache.SourceType 属性
    /// </summary>
    XlPivotTableSourceType SourceType { get; }

    /// <summary>
    /// 获取数据透视表缓存的源数据
    /// 对应 PivotCache.SourceData 属性
    /// 可以是 string, Range, ListObject 等
    /// </summary>
    object SourceData { get; }


    /// <summary>
    /// 获取数据透视表缓存的记录数
    /// 对应 PivotCache.RecordCount 属性
    /// </summary>
    int RecordCount { get; }

    /// <summary>
    /// 获取数据透视表缓存的版本
    /// 对应 PivotCache.Version 属性
    /// </summary>
    XlPivotTableVersionList Version { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 刷新数据透视表缓存
    /// 对应 PivotCache.Refresh 方法
    /// </summary>
    void Refresh();

    /// <summary>
    /// 将指定数据透视表的刷新计时器重置为使用 RefreshPeriod 属性设置的最后一个间隔。
    /// </summary>
    void ResetTimer();

    /// <summary>
    /// 为指定的数据透视表缓存建立连接。
    /// </summary>
    void MakeConnection();

    /// <summary>
    /// 从 PivotCache 对象创建独立数据透视图。 返回<see cref="IExcelShape"/>对象。
    /// </summary>
    /// <param name="chartDestination">“目标”工作表</param>
    /// <param name="xlChartType">图表的类型</param>
    /// <param name="left">从对象左边界至 A 列左边界（在工作表上）或图表区左边界（在图表上）的距离，以磅为单位。</param>
    /// <param name="top">从图形区域中最上端的图形的顶端到工作表顶端的距离，以磅为单位。</param>
    /// <param name="width">对象的宽度，以磅为单位。</param>
    /// <param name="height">对象的高度，以磅为单位。</param>
    /// <returns></returns>
    IExcelShape? CreatePivotChart(IExcelWorksheet chartDestination, XlChartType? xlChartType,
        int? left, int? top, int? width, int? height);

    /// <summary>
    /// 基于 PivotCache 对象创建数据透视表。
    /// </summary>
    /// <param name="tableDestination">必需的 对象。 数据透视表目标区域（工作表中用于放置所生成的数据透视表的区域）左上角的单元格。 目标区域必须位于包含表达式指定的 PivotCache 对象的工作簿中的工作表上。</param>
    /// <param name="tableName">可选 对象。 新的数据透视表的名称。</param>
    /// <param name="readData">可选 对象。 如果该值为 True，则创建一个包含外部数据库中所有记录的数据透视表高速缓存；此高速缓存可以很大。 如果为 False，则允许在实际读取数据之前将某些字段设置为基于服务器的页字段。</param>
    /// <param name="defaultVersion">可选 对象。 数据透视表的默认版本。</param>
    /// <returns></returns>
    IExcelPivotTable? CreatePivotTable(IExcelRange tableDestination, string? tableName = null,
        bool? readData = null, XlPivotTableVersionList? defaultVersion = null);
    #endregion
}
