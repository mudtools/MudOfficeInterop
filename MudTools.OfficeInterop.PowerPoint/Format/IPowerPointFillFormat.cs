//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 填充格式接口
/// </summary>
public interface IPowerPointFillFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    int ForeColor { get; set; }

    /// <summary>
    /// 获取或设置背景色
    /// </summary>
    int BackColor { get; set; }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取填充类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取或设置渐变颜色类型
    /// </summary>
    int GradientColorType { get; }

    /// <summary>
    /// 获取或设置渐变样式
    /// </summary>
    int GradientStyle { get; }

    /// <summary>
    /// 获取或设置渐变变体
    /// </summary>
    int GradientVariant { get; }

    /// <summary>
    /// 获取或设置图案类型
    /// </summary>
    int Pattern { get; }

    /// <summary>
    /// 获取或设置纹理类型
    /// </summary>
    int TextureType { get; set; }

    /// <summary>
    /// 获取或设置纹理名称
    /// </summary>
    string TextureName { get; set; }

    /// <summary>
    /// 设置纯色填充
    /// </summary>
    void Solid();

    /// <summary>
    /// 设置图案填充
    /// </summary>
    /// <param name="pattern">图案类型</param>
    void Patterned(int pattern);

    /// <summary>
    /// 设置渐变填充
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    /// <param name="presetGradientType">预设渐变类型</param>
    void Gradient(int style, int variant, int presetGradientType);

    /// <summary>
    /// 设置纹理填充
    /// </summary>
    /// <param name="textureFile">纹理文件路径</param>
    /// <param name="textureType">纹理类型</param>
    void Textured(string textureFile, MsoPresetTexture textureType = MsoPresetTexture.msoPresetTextureMixed);

    /// <summary>
    /// 设置图片填充
    /// </summary>
    /// <param name="pictureFile">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    void UserPicture(string pictureFile, bool linkToFile = false, bool saveWithDocument = true);

    /// <summary>
    /// 设置预设纹理填充
    /// </summary>
    /// <param name="presetTexture">预设纹理类型</param>
    void PresetTextured(MsoPresetTexture presetTexture);

    /// <summary>
    /// 设置预设渐变填充
    /// </summary>
    /// <param name="presetGradientType">预设渐变类型</param>
    void PresetGradient(MsoPresetGradientType presetGradientType);

    /// <summary>
    /// 设置自定义颜色渐变
    /// </summary>
    /// <param name="color1">起始颜色</param>
    /// <param name="color2">结束颜色</param>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    void TwoColorGradient(int color1, int color2, int style, int variant);

    /// <summary>
    /// 重置填充格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制填充格式
    /// </summary>
    /// <returns>复制的填充格式对象</returns>
    IPowerPointFillFormat Duplicate();

    /// <summary>
    /// 应用填充格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    void ApplyTo(IPowerPointShape shape);


    /// <summary>
    /// 设置填充颜色
    /// </summary>
    /// <param name="foregroundColor">前景色</param>
    /// <param name="backgroundColor">背景色</param>
    void SetColors(int foregroundColor, int backgroundColor = 0);

    /// <summary>
    /// 获取填充信息
    /// </summary>
    /// <returns>填充信息字符串</returns>
    string GetFillInfo();
}
