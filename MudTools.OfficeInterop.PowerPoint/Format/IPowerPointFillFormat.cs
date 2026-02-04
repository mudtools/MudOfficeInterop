//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示对象的填充格式。
/// 提供设置和获取填充属性（如颜色、渐变、纹理等）的功能。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointFillFormat : IOfficeObject<IPowerPointFillFormat, MsPowerPoint.FillFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此填充格式的应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此填充格式的应用程序的创建者代码。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取此填充格式的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 将填充设置为背景填充。
    /// </summary>
    void Background();

    /// <summary>
    /// 设置单色渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变体，取值范围为1到4。</param>
    /// <param name="degree">渐变角度，取值范围为0.0到1.0。</param>
    void OneColorGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant, float degree);

    /// <summary>
    /// 设置图案填充。
    /// </summary>
    /// <param name="pattern">图案类型。</param>
    void Patterned([ComNamespace("MsCore")] MsoPatternType pattern);

    /// <summary>
    /// 设置预设渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变体，取值范围为1到4。</param>
    /// <param name="presetGradientType">预设渐变类型。</param>
    void PresetGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant, [ComNamespace("MsCore")] MsoPresetGradientType presetGradientType);

    /// <summary>
    /// 设置预设纹理填充。
    /// </summary>
    /// <param name="presetTexture">预设纹理类型。</param>
    void PresetTextured([ComNamespace("MsCore")] MsoPresetTexture presetTexture);

    /// <summary>
    /// 设置纯色填充。
    /// </summary>
    void Solid();

    /// <summary>
    /// 设置双色渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变体，取值范围为1到4。</param>
    void TwoColorGradient([ComNamespace("MsCore")] MsoGradientStyle style, int variant);

    /// <summary>
    /// 使用指定图片文件设置填充。
    /// </summary>
    /// <param name="pictureFile">图片文件路径。</param>
    void UserPicture(string pictureFile);

    /// <summary>
    /// 使用指定纹理文件设置填充。
    /// </summary>
    /// <param name="textureFile">纹理文件路径。</param>
    void UserTextured(string textureFile);

    /// <summary>
    /// 获取或设置填充的背景颜色。
    /// </summary>
    IPowerPointColorFormat? BackColor { get; set; }

    /// <summary>
    /// 获取或设置填充的前景颜色。
    /// </summary>
    IPowerPointColorFormat? ForeColor { get; set; }

    /// <summary>
    /// 获取渐变填充的颜色类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoGradientColorType GradientColorType { get; }

    /// <summary>
    /// 获取单色渐变的渐变程度（0.0到1.0）。
    /// </summary>
    float GradientDegree { get; }

    /// <summary>
    /// 获取渐变填充的样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoGradientStyle GradientStyle { get; }

    /// <summary>
    /// 获取渐变填充的变体（1到4）。
    /// </summary>
    int GradientVariant { get; }

    /// <summary>
    /// 获取图案填充的图案类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPatternType Pattern { get; }

    /// <summary>
    /// 获取预设渐变填充的类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetGradientType PresetGradientType { get; }

    /// <summary>
    /// 获取预设纹理填充的纹理类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTexture PresetTexture { get; }

    /// <summary>
    /// 获取纹理填充的纹理名称。
    /// </summary>
    string TextureName { get; }

    /// <summary>
    /// 获取纹理填充的类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextureType TextureType { get; }

    /// <summary>
    /// 获取或设置填充的透明度（0.0到1.0）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取填充类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoFillType Type { get; }

    /// <summary>
    /// 获取或设置填充是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取渐变填充的渐变停止点集合。
    /// </summary>
    IOfficeGradientStops? GradientStops { get; }

    /// <summary>
    /// 获取或设置纹理填充的水平偏移量。
    /// </summary>
    float TextureOffsetX { get; set; }

    /// <summary>
    /// 获取或设置纹理填充的垂直偏移量。
    /// </summary>
    float TextureOffsetY { get; set; }

    /// <summary>
    /// 获取或设置纹理填充的对齐方式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextureAlignment TextureAlignment { get; set; }

    /// <summary>
    /// 获取或设置纹理填充的水平缩放比例。
    /// </summary>
    float TextureHorizontalScale { get; set; }

    /// <summary>
    /// 获取或设置纹理填充的垂直缩放比例。
    /// </summary>
    float TextureVerticalScale { get; set; }

    /// <summary>
    /// 获取或设置纹理是否平铺。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool TextureTile { get; set; }

    /// <summary>
    /// 获取或设置填充是否随对象旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotateWithObject { get; set; }

    /// <summary>
    /// 获取图片效果集合。
    /// </summary>
    IOfficePictureEffects? PictureEffects { get; }

    /// <summary>
    /// 获取或设置渐变填充的角度。
    /// </summary>
    float GradientAngle { get; set; }
}