//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示迷你图中数据点标记的颜色。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSparkColor : IOfficeObject<IExcelSparkColor, MsExcel.SparkColor>, IDisposable
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
    /// 获取或设置一个值，指示该点是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取一个 FormatColor 对象，可用于设置迷你图中数据点标记的颜色。
    /// </summary>
    IExcelFormatColor? Color { get; }
}