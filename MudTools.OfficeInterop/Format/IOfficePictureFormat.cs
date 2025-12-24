//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中图片格式的接口封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficePictureFormat : IDisposable
{
    /// <summary>
    /// 获取图片的裁剪信息，如果图片未被裁剪则返回null。
    /// </summary>
    IOfficeCrop? Crop { get; }

    /// <summary>
    /// 获取或设置图片的亮度（-1.0 到 1.0 之间）。
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图片的对比度（0.0 到 1.0 之间）。
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置图片的透明度颜色。
    /// </summary>
    int TransparencyColor { get; set; }

    /// <summary>
    /// 获取或设置图片的裁剪区域。
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置图片的右边裁剪区域。
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置图片的顶部裁剪区域。
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置图片的底部裁剪区域。
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取或设置图片是否使用透明度颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 按指定增量调整图片亮度。
    /// </summary>
    /// <param name="increment">亮度调整的增量值，正值增加亮度，负值减少亮度。</param>
    void IncrementBrightness(float increment);

    /// <summary>
    /// 按指定增量调整图片对比度。
    /// </summary>
    /// <param name="increment">对比度调整的增量值，正值增加对比度，负值减少对比度。</param>
    void IncrementContrast(float increment);
}