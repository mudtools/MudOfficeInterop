//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中图形对象的填充格式（Fill Format）的抽象接口。
/// 封装了 Microsoft.Office.Interop.Word.FillFormat 的常用功能，便于测试和解耦。
/// </summary>
public interface IWordFillFormat : IDisposable
{
    #region 属性

    /// <summary>
    /// 获取此填充格式所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此填充格式的父对象（通常是 Shape 或 InlineShape）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置填充的前景色（例如纯色填充或图案前景色）。
    /// </summary>
    IWordColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取或设置填充的背景色（主要用于图案填充）。
    /// </summary>
    IWordColorFormat? BackColor { get; }

    /// <summary>
    /// 获取或设置填充的透明度，取值范围为 0.0（完全不透明）到 1.0（完全透明）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置填充是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取当前填充的类型（如纯色、渐变、图案、纹理、图片等）。
    /// </summary>
    MsoFillType Type { get; }

    /// <summary>
    /// 获取渐变填充的颜色类型（单色、双色等）。
    /// </summary>
    MsoGradientColorType GradientColorType { get; }

    /// <summary>
    /// 获取渐变填充的方向样式（水平、垂直、对角线等）。
    /// </summary>
    MsoGradientStyle GradientStyle { get; }

    /// <summary>
    /// 获取或设置渐变填充的角度（以度为单位）。
    /// </summary>
    float GradientAngle { get; set; }

    /// <summary>
    /// 获取当前使用的图案类型（仅当填充类型为图案时有效）。
    /// </summary>
    MsoPatternType Pattern { get; }

    /// <summary>
    /// 获取预设纹理类型（仅当填充类型为预设纹理时有效）。
    /// </summary>
    MsoPresetTexture PresetTexture { get; }

    /// <summary>
    /// 获取用户自定义纹理的文件名（仅当使用自定义纹理时有效）。
    /// </summary>
    string TextureName { get; }

    /// <summary>
    /// 获取当前纹理类型（预设或自定义）。
    /// </summary>
    MsoTextureType TextureType { get; }

    /// <summary>
    /// 获取预设渐变类型（如“日出”、“金属”等）。
    /// </summary>
    MsoPresetGradientType PresetGradientType { get; }

    /// <summary>
    /// 获取渐变停靠点（Gradient Stops）的数量。
    /// </summary>
    int GradientStopsCount { get; }

    /// <summary>
    /// 获取或设置纹理在 X 轴上的偏移量（以磅为单位）。
    /// </summary>
    float TextureOffsetX { get; set; }

    /// <summary>
    /// 获取或设置纹理在 Y 轴上的偏移量（以磅为单位）。
    /// </summary>
    float TextureOffsetY { get; set; }

    /// <summary>
    /// 获取或设置纹理的对齐方式。
    /// </summary>
    MsoTextureAlignment TextureAlignment { get; set; }

    /// <summary>
    /// 获取或设置纹理在水平方向上的缩放比例（1.0 表示 100%）。
    /// </summary>
    float TextureHorizontalScale { get; set; }

    /// <summary>
    /// 获取或设置纹理在垂直方向上的缩放比例（1.0 表示 100%）。
    /// </summary>
    float TextureVerticalScale { get; set; }

    /// <summary>
    /// 获取或设置是否平铺纹理以填充整个区域。
    /// </summary>
    bool TextureTile { get; set; }

    /// <summary>
    /// 获取或设置填充是否随对象旋转而旋转。
    /// </summary>
    bool RotateWithObject { get; set; }

    /// <summary>
    /// 获取应用于填充的图片效果集合（如阴影、发光、柔化边缘等）。
    /// </summary>
    IOfficePictureEffects? PictureEffects { get; }

    #endregion

    #region 方法

    /// <summary>
    /// 将填充设置为纯色（默认颜色）。
    /// </summary>
    void Solid();

    /// <summary>
    /// 将填充设置为指定 RGB 颜色的纯色。
    /// </summary>
    /// <param name="color">RGB 颜色值（例如 0xFF0000 表示红色）。</param>
    void Solid(int color);

    /// <summary>
    /// 应用预设的渐变填充效果。
    /// </summary>
    /// <param name="style">渐变方向样式。</param>
    /// <param name="variant">渐变变体（1-4）。</param>
    /// <param name="presetGradientType">预设渐变类型。</param>
    void PresetGradient(MsoGradientStyle style, int variant, MsoPresetGradientType presetGradientType);

    /// <summary>
    /// 应用预设纹理填充。
    /// </summary>
    /// <param name="presetTexture">预设纹理类型。</param>
    void PresetTextured(MsoPresetTexture presetTexture);

    /// <summary>
    /// 应用图案填充，并指定前景色和背景色。
    /// </summary>
    /// <param name="pattern">图案类型。</param>
    /// <param name="foregroundColor">前景色的 RGB 值。</param>
    /// <param name="backgroundColor">背景色的 RGB 值。</param>
    void Patterned(MsoPatternType pattern, int foregroundColor, int backgroundColor);

    /// <summary>
    /// 使用指定的图片文件作为填充。
    /// </summary>
    /// <param name="pictureFile">图片文件的完整路径。</param>
    void UserPicture(string pictureFile);

    /// <summary>
    /// 使用指定的纹理图片作为填充。
    /// </summary>
    /// <param name="textureFile">纹理图片文件的完整路径。</param>
    void UserTextured(string textureFile);

    /// <summary>
    /// 设置填充的透明度。
    /// </summary>
    /// <param name="transparency">透明度值（0.0 到 1.0）。</param>
    void SetTransparent(float transparency);

    /// <summary>
    /// 清除填充（使其不可见）。
    /// </summary>
    void Clear();

    /// <summary>
    /// 将当前填充格式复制到目标填充对象。
    /// </summary>
    /// <param name="targetFill">目标填充对象。</param>
    /// <exception cref="InvalidOperationException">当 COM 操作失败时抛出。</exception>
    void CopyTo(IWordFillFormat targetFill);

    #endregion
}