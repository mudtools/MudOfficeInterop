using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ColorFormat 的接口，用于操作颜色格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordColorFormat : IOfficeObject<IWordColorFormat, MsWord.ColorFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置颜色的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置是否覆盖打印设置。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool OverPrint { get; set; }

    /// <summary>
    /// 获取或设置颜色的RGB值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color RGB { get; set; }

    /// <summary>
    /// 获取或设置颜色亮度。
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取颜色格式的类型。
    /// </summary>
    MsoColorType Type { get; }

    /// <summary>
    /// 获取或设置方案颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color SchemeColor { get; set; }

    /// <summary>
    /// 获取或设置对象的主题颜色索引。
    /// </summary>
    WdThemeColorIndex ObjectThemeColor { get; set; }

    /// <summary>
    /// 获取或设置颜色的深浅和阴影。
    /// </summary>
    float TintAndShade { get; set; }

    /// <summary>
    /// 获取或设置颜色中的黑色成分值（CMYK颜色模型中的K值）。
    /// </summary>
    int Black { get; set; }

    /// <summary>
    /// 获取或设置颜色中的黄色成分值（CMYK颜色模型中的Y值）。
    /// </summary>
    int Yellow { get; set; }

    /// <summary>
    /// 获取或设置颜色中的品红色成分值（CMYK颜色模型中的M值）。
    /// </summary>
    int Magenta { get; set; }

    /// <summary>
    /// 获取或设置颜色中的青色成分值（CMYK颜色模型中的C值）。
    /// </summary>
    int Cyan { get; set; }

    /// <summary>
    /// 获取或设置颜色的墨水浓度。
    /// </summary>
    [MethodIndex]
    float? Ink(int index);


    /// <summary>
    /// 设置颜色的CMYK值。
    /// </summary>
    /// <param name="Cyan">青色值</param>
    /// <param name="Magenta">品红色值</param>
    /// <param name="Yellow">黄色值</param>
    /// <param name="Black">黑色值</param>
    void SetCMYK(int Cyan, int Magenta, int Yellow, int Black);
}