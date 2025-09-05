namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ColorFormat 的接口，用于操作颜色格式。
/// </summary>
public interface IWordColorFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置颜色的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置是否覆盖打印设置。
    /// </summary>
    bool OverPrint { get; set; }

    /// <summary>
    /// 获取或设置颜色的RGB值。
    /// </summary>
    int RGB { get; set; }

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
    int SchemeColor { get; set; }

    /// <summary>
    /// 获取或设置对象的主题颜色索引。
    /// </summary>
    WdThemeColorIndex ObjectThemeColor { get; set; }

    /// <summary>
    /// 获取或设置颜色的深浅和阴影。
    /// </summary>
    float TintAndShade { get; set; }

    /// <summary>
    /// 设置颜色的CMYK值。
    /// </summary>
    /// <param name="Cyan">青色值</param>
    /// <param name="Magenta">品红色值</param>
    /// <param name="Yellow">黄色值</param>
    /// <param name="Black">黑色值</param>
    void SetCMYK(int Cyan, int Magenta, int Yellow, int Black);
}