//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Axes 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Axes 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelAxes : IEnumerable<IExcelAxis?>, IOfficeObject<IExcelAxes, MsExcel.Axes>, IDisposable
{

    /// <summary>
    /// 获取坐标轴集合所在的父对象（通常是 Chart）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取坐标轴集合所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    /// <summary>
    /// 获取坐标轴集合中的坐标轴数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的坐标轴对象
    /// </summary>
    /// <returns>坐标轴对象</returns>
    IExcelAxis? this[XlAxisType type, XlAxisGroup axisGroup = XlAxisGroup.xlPrimary] { get; }

    /// <summary> 
    /// 获取指定索引的坐标轴对象
    /// 索引从1开始
    /// </summary>
    /// <param name="type">坐标轴类型</param>
    /// <returns>坐标轴对象</returns>
    IExcelAxis? this[XlAxisType type] { get; }


}
