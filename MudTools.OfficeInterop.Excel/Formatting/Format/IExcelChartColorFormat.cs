//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

public interface IExcelChartColorFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取颜色类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取或设置RGB颜色值
    /// </summary>
    int RGB { get; }

    /// <summary>
    /// 获取或设置颜色方案索引
    /// </summary>
    int SchemeColor { get; set; }

    /// <summary>
    /// 获取颜色的HSL值
    /// </summary>
    /// <param name="hue">色相</param>
    /// <param name="saturation">饱和度</param>
    /// <param name="lightness">亮度</param>
    void GetHSL(out double hue, out double saturation, out double lightness);

    /// <summary>
    /// 混合两种颜色
    /// </summary>
    /// <param name="color1">第一种颜色</param>
    /// <param name="color2">第二种颜色</param>
    /// <param name="ratio">混合比例</param>
    /// <returns>混合后的颜色</returns>
    int BlendColors(int color1, int color2, double ratio);
}