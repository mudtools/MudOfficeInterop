//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop.Excel;
public interface IExcelColorFormat : IDisposable
{
    /// <summary>
    /// 获取颜色格式的父级对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置颜色类型
    /// </summary>
    MsoColorType Type { get; }

    /// <summary>
    /// 获取或设置RGB颜色值
    /// </summary>
    int RGB { get; set; }

    /// <summary>
    /// 获取或设置颜色的应用程序版本
    /// </summary>
    IExcelApplication Application { get; }


    /// <summary>
    /// 获取颜色的十六进制表示
    /// </summary>
    /// <returns>十六进制颜色字符串</returns>
    string ToHexString();

    /// <summary>
    /// 获取颜色的HSL值
    /// </summary>
    /// <param name="hue">色相</param>
    /// <param name="saturation">饱和度</param>
    /// <param name="lightness">亮度</param>
    void GetHSL(out double hue, out double saturation, out double lightness);

    /// <summary>
    /// 设置HSL颜色值
    /// </summary>
    /// <param name="hue">色相</param>
    /// <param name="saturation">饱和度</param>
    /// <param name="lightness">亮度</param>
    void SetHSL(double hue, double saturation, double lightness);

    /// <summary>
    /// 获取颜色名称
    /// </summary>
    /// <returns>颜色名称</returns>
    string GetColorName();


    /// <summary>
    /// 获取颜色的对比色
    /// </summary>
    /// <returns>对比色</returns>
    int GetContrastColor();

    /// <summary>
    /// 混合两种颜色
    /// </summary>
    /// <param name="color1">第一种颜色</param>
    /// <param name="color2">第二种颜色</param>
    /// <param name="ratio">混合比例</param>
    /// <returns>混合后的颜色</returns>
    int BlendColors(int color1, int color2, double ratio);

    /// <summary>
    /// 获取颜色亮度值
    /// </summary>
    /// <returns>亮度值</returns>
    double GetLuminance();

    /// <summary>
    /// 获取颜色饱和度值
    /// </summary>
    /// <returns>饱和度值</returns>
    double GetSaturation();
}