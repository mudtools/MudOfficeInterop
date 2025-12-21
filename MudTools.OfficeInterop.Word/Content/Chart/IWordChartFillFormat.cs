//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表填充格式的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartFillFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置填充前景色。
    /// </summary>
    IWordChartColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取或设置填充背景色。
    /// </summary>
    IWordChartColorFormat? BackColor { get; }

    /// <summary>
    /// 获取或设置填充样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPatternType Pattern { get; }

    /// <summary>
    /// 获取渐变颜色类型。
    /// </summary>
    MsoGradientColorType GradientColorType { get; }

    /// <summary>
    /// 获取渐变度数。
    /// </summary>
    float GradientDegree { get; }

    /// <summary>
    /// 获取渐变样式。
    /// </summary>
    MsoGradientStyle GradientStyle { get; }

    /// <summary>
    /// 获取渐变变体。
    /// </summary>
    int GradientVariant { get; }

    /// <summary>
    /// 获取预设渐变类型。
    /// </summary>
    MsoPresetGradientType PresetGradientType { get; }

    /// <summary>
    /// 获取预设纹理。
    /// </summary>
    MsoPresetTexture PresetTexture { get; }

    /// <summary>
    /// 获取纹理名称。
    /// </summary>
    string TextureName { get; }

    /// <summary>
    /// 获取纹理类型。
    /// </summary>
    MsoTextureType TextureType { get; }

    /// <summary>
    /// 获取填充是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; }

    /// <summary>
    /// 获取或设置填充类型。
    /// </summary>
    MsoFillType Type { get; }

    /// <summary>
    /// 设置渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变量。 取值范围为 1 到 4，分别与“填充效果”对话框中“渐变”选项卡上的四个变量相对应。 如果 GradientStyle 为 msoGradientFromCenter，则 Variant 参数只能为 1 或 2。</param>
    /// <param name="degree">渐变角度。可以为 0.0（暗）到 1.0（亮）之间的值。</param>
    void OneColorGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant, float degree);

    /// <summary>
    /// 设置双色渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变量。 取值范围为 1 到 4，分别与“填充效果”对话框中“渐变”选项卡上的四个变量相对应。 如果 GradientStyle 为 msoGradientFromCenter，则 Variant 参数只能为 1 或 2。</param>
    void TwoColorGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant);

    /// <summary>
    /// 设置纯色填充。
    /// </summary>
    void Solid();

    /// <summary>
    /// 使用用户自定义图片作为填充。
    /// </summary>
    /// <param name="pictureFile">图片文件路径。</param>
    /// <param name="pictureFormat">图片格式类型。</param>
    /// <param name="pictureStackUnit">图片堆叠单位。</param>
    /// <param name="picturePlacement">图片放置位置。</param>
    void UserPicture(string? pictureFile, XlChartPictureType? pictureFormat = null, double? pictureStackUnit = null, XlChartPicturePlacement? picturePlacement = null);

    /// <summary>
    /// 使用用户自定义纹理作为填充。
    /// </summary>
    /// <param name="textureFile">纹理文件路径。</param>
    void UserTextured(string textureFile);

    /// <summary>
    /// 设置预设渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变量。 取值范围为 1 到 4，分别与“填充效果”对话框中“渐变”选项卡上的四个变量相对应。 如果 GradientStyle 为 msoGradientFromCenter，则 Variant 参数只能为 1 或 2。</param>
    /// <param name="presetGradientType">预设渐变类型。</param>
    void PresetGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant, [ComNamespace("MsCore")] MsoPresetGradientType presetGradientType);


    /// <summary>
    /// 设置纹理填充。
    /// </summary>
    /// <param name="texture">纹理类型。</param>
    void PresetTextured([ComNamespace("MsCore")] MsoPresetTexture texture);
}