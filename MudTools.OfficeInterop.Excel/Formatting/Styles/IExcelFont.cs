//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 定义Excel字体样式的基本接口
/// 该接口提供了操作Excel单元格字体的各种属性和方法
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelFont : IOfficeObject<IExcelFont, MsExcel.Font>, IDisposable
{

    /// <summary>
    /// 获取字体的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取字体所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    double Size { get; set; }

    /// <summary>
    /// 获取或设置是否粗体
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置字体背景
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlBackground Background { get; set; }

    /// <summary>
    /// 获取或设置是否删除线
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Strikethrough { get; set; }

    /// <summary>
    /// 获取或设置字体样式
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string FontStyle { get; set; }

    /// <summary>
    /// 获取或设置字体颜色索引
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置字体颜色（RGB值）
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置下划线样式
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlUnderlineStyle Underline { get; set; }

    /// <summary>
    /// 获取或设置是否为上标
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置是否为下标
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Subscript { get; set; }

    /// <summary>
    /// 字体为空心字体则该属性值为 True。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlThemeFont OutlineFont { get; set; }

    /// <summary>
    /// 指定对象关联的应用字体方案中的主题字体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlThemeFont ThemeFont { get; set; }

    /// <summary>
    /// 可以为属性输入从 -1 (最暗) 到 1 (最亮) TintAndShade 的数字。 零 (0) 为中性。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float TintAndShade { get; set; }
}
