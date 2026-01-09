//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示迷你图上数据点的标记设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSparkPoints : IOfficeObject<IExcelSparkPoints, MsExcel.SparkPoints>, IDisposable
{
    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上负数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Negative { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上正数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Markers { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上最高数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Highpoint { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上最低数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Lowpoint { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上第一个数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Firstpoint { get; }

    /// <summary>
    /// 获取一个 SparkColor 对象，表示迷你图上最后一个数据点的标记颜色和可见性。
    /// </summary>
    IExcelSparkColor? Lastpoint { get; }
}