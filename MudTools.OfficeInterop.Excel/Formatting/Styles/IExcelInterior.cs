//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Interior 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Interior 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelInterior : IOfficeObject<IExcelInterior, MsExcel.Interior>, IDisposable
{

    /// <summary>
    /// 获取内部对象所在的父对象
    /// 对应 Interior.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取内部对象所在的Application对象
    /// 对应 Interior.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #region 基础属性

    /// <summary>
    /// 获取或设置内部颜色
    /// 对应 Interior.Color 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置内部图案类型
    /// 对应 Interior.Pattern 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlPattern Pattern { get; set; }

    /// <summary>
    /// 获取或设置内部颜色索引
    /// 对应 Interior.ColorIndex 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置图案颜色
    /// 对应 Interior.PatternColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color PatternColor { get; set; }

    /// <summary>
    /// 获取或设置主题颜色
    /// 对应 Interior.ThemeColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color ThemeColor { get; set; }

    /// <summary>
    /// 获取或设置着色和阴影
    /// 对应 Interior.TintAndShade 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    double TintAndShade { get; set; }

    /// <summary>
    /// 获取或设置图案主题颜色
    /// 对应 Interior.PatternThemeColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color PatternThemeColor { get; set; }

    /// <summary>
    /// 获取或设置图案颜色索引
    /// 对应 Interior.PatternColorIndex 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex PatternColorIndex { get; set; }

    /// <summary>
    /// 获取或设置图案着色和阴影
    /// 对应 Interior.PatternTintAndShade 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    double PatternTintAndShade { get; set; }
    #endregion
}