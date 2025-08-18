//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 背景接口
/// </summary>
public interface IPowerPointBackground : IDisposable
{
    /// <summary>
    /// 获取填充格式
    /// </summary>
    IPowerPointFillFormat Fill { get; }

    /// <summary>
    /// 获取颜色方案
    /// </summary>
    IPowerPointColorScheme ColorScheme { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置背景样式
    /// </summary>
    int Style { get; set; }

    /// <summary>
    /// 获取或设置背景类型
    /// </summary>
    int Type { get; }

    /// <summary>
    /// 获取或设置是否显示背景图形
    /// </summary>
    bool DisplayBackground { get; set; }

    /// <summary>
    /// 应用纯色背景
    /// </summary>
    /// <param name="color">颜色</param>
    void ApplySolidBackground(int color);

    /// <summary>
    /// 应用渐变背景
    /// </summary>
    /// <param name="style">渐变样式</param>
    /// <param name="variant">渐变变体</param>
    /// <param name="color1">起始颜色</param>
    /// <param name="color2">结束颜色</param>
    void ApplyGradientBackground(int style, int variant, int color1, int color2);

    /// <summary>
    /// 应用图片背景
    /// </summary>
    /// <param name="pictureFile">图片文件路径</param>
    /// <param name="tile">是否平铺</param>
    void ApplyPictureBackground(string pictureFile, bool tile = false);

    /// <summary>
    /// 应用纹理背景
    /// </summary>
    /// <param name="textureType">纹理类型</param>
    void ApplyTextureBackground(MsoPresetTexture textureType);

    /// <summary>
    /// 应用主题背景
    /// </summary>
    /// <param name="themeIndex">主题索引</param>
    void ApplyThemeBackground(int themeIndex);

    /// <summary>
    /// 重置背景
    /// </summary>
    void Reset();

    /// <summary>
    /// 应用到所有幻灯片
    /// </summary>
    void ApplyToAll();

    /// <summary>
    /// 获取背景信息
    /// </summary>
    /// <returns>背景信息字符串</returns>
    string GetBackgroundInfo();
}