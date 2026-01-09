//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ChartColorFormat 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartColorFormat 的安全访问和操作
/// 用于设置形状或图表元素的颜色
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelChartColorFormat : IOfficeObject<IExcelChartColorFormat, MsExcel.ChartColorFormat>, IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取颜色类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取或设置RGB颜色值
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color RGB { get; }

    /// <summary>
    /// 获取或设置颜色方案索引
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color SchemeColor { get; set; }
}