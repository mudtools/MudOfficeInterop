//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 定义Excel颜色格式的接口，提供对Excel中颜色相关属性和操作的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelColorFormat : IOfficeObject<IExcelColorFormat, MsExcel.ColorFormat>, IDisposable
{
    /// <summary>
    /// 获取颜色格式的父级对象
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取或设置颜色的应用程序版本
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置颜色类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoColorType Type { get; }

    /// <summary>
    /// 获取或设置对象的主题颜色索引
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoThemeColorIndex ObjectThemeColor { get; set; }

    /// <summary>
    /// 获取或设置颜色的亮度值
    /// </summary>
    float? Brightness { get; set; }

    /// <summary>
    /// 获取或设置颜色的色调和阴影值
    /// </summary>
    float? TintAndShade { get; set; }

    /// <summary>
    /// 获取或设置RGB颜色值
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color RGB { get; set; }

    /// <summary>
    /// 获取或设置颜色方案索引
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color SchemeColor { get; set; }


}