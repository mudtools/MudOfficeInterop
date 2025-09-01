//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureFormat 的接口，用于操作图片格式。
/// </summary>
public interface IWordPictureFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    IWordCrop? Crop { get; }
    /// <summary>
    /// 获取或设置图片的亮度（-1.0到1.0之间）。
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图片的对比度（0.0到1.0之间）。
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置图片的颜色类型。
    /// </summary>
    MsoPictureColorType ColorType { get; set; }

    /// <summary>
    /// 获取或设置图片的裁剪左边缘（磅）。
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置图片的裁剪右边缘（磅）。
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置图片的裁剪上边缘（磅）。
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置图片的裁剪下边缘（磅）。
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取或设置图片的透明色。
    /// </summary>
    int TransparencyColor { get; set; }

    /// <summary>
    /// 获取或设置图片是否透明。
    /// </summary>
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 获取或设置图片的柔化边缘格式。
    /// </summary>
    IWordSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取或设置图片的光泽格式。
    /// </summary>
    IWordReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取或设置图片的反射格式。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取图片是否为链接图片。
    /// </summary>
    bool IsLinked { get; }

    /// <summary>
    /// 获取图片的文件名。
    /// </summary>
    string Filename { get; }

    /// <summary>
    /// 获取图片的文件大小（字节）。
    /// </summary>
    long FileSize { get; }

    /// <summary>
    /// 调整图片亮度。
    /// </summary>
    /// <param name="brightness">亮度值（-1.0到1.0）。</param>
    void AdjustBrightness(float brightness);

    /// <summary>
    /// 调整图片对比度。
    /// </summary>
    /// <param name="contrast">对比度值（0.0到1.0）。</param>
    void AdjustContrast(float contrast);

    /// <summary>
    /// 重置图片格式为原始状态。
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置透明色。
    /// </summary>
    /// <param name="rgb">RGB颜色值。</param>
    void SetTransparentColor(int rgb);

    /// <summary>
    /// 复制图片格式到另一个对象。
    /// </summary>
    /// <param name="targetPicture">目标图片格式对象。</param>
    void CopyTo(IWordPictureFormat targetPicture);

    /// <summary>
    /// 更新链接的图片。
    /// </summary>
    /// <returns>是否更新成功。</returns>
    bool Update();

    /// <summary>
    /// 断开图片链接。
    /// </summary>
    /// <returns>是否断开成功。</returns>
    bool BreakLink();

    /// <summary>
    /// 验证图片参数是否有效。
    /// </summary>
    /// <param name="brightness">亮度值。</param>
    /// <param name="contrast">对比度值。</param>
    /// <returns>参数是否有效。</returns>
    bool ValidateParameters(float brightness, float contrast);

    /// <summary>
    /// 获取图片是否为透明图片。
    /// </summary>
    bool HasTransparency { get; }

    /// <summary>
    /// 获取图片是否为灰度模式。
    /// </summary>
    bool IsGrayscale { get; }
}