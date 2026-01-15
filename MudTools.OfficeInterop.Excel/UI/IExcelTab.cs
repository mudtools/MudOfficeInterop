//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示图表或工作表中的选项卡。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTab : IOfficeObject<IExcelTab, MsExcel.Tab>, IDisposable
{
    /// <summary>
    /// 获取父级工作表
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置选项卡的主要颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置指定图表选项卡的颜色。
    /// </summary>
    XlColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置与指定对象关联的应用配色方案中的主题颜色。可读写。
    /// </summary>
    XlThemeColor ThemeColor { get; set; }

    /// <summary>
    /// 获取或设置用于调亮或调暗颜色的单精度值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float? TintAndShade { get; set; }
}