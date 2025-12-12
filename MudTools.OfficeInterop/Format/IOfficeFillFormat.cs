//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office填充格式对象的接口
/// 封装了Microsoft.Office.Core.FillFormat COM对象
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeFillFormat : IDisposable
{
    /// <summary>
    /// 获取前景颜色格式
    /// </summary>
    IOfficeColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取背景颜色格式
    /// </summary>
    IOfficeColorFormat? BackColor { get; }

    /// <summary>
    /// 获取或设置填充透明度（0-1之间，0为完全不透明，1为完全透明）
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取填充类型
    /// </summary>
    MsoFillType Type { get; }

    /// <summary>
    /// 获取或设置渐变角度
    /// </summary>
    float GradientAngle { get; set; }

    /// <summary>
    /// 设置为单色渐变填充
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    /// <param name="degree">渐变程度</param>
    void OneColorGradient(MsoGradientStyle style, int variant, float degree);

    /// <summary>
    /// 设置为双色渐变填充
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    void TwoColorGradient(MsoGradientStyle style, int variant);

    /// <summary>
    /// 设置为图案填充
    /// </summary>
    /// <param name="pattern">图案类型</param>
    void Patterned(MsoPatternType pattern);

    /// <summary>
    /// 设置为图片填充
    /// </summary>
    /// <param name="imagePath">图片路径</param>
    void UserPicture(string imagePath);

    /// <summary>
    /// 设置为实心填充
    /// </summary>
    void Solid();

    /// <summary>
    /// 设置为纹理填充
    /// </summary>
    void UserTextured(string textureFile);

    /// <summary>
    /// 重置为无填充
    /// </summary>
    void Background();

    /// <summary>
    /// 获取纹理名称
    /// </summary>
    string TextureName { get; }

    /// <summary>
    /// 获取纹理类型
    /// </summary>
    MsoTextureType TextureType { get; }
}