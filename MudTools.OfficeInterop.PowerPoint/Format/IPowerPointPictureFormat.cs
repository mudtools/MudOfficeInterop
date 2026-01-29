//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 中图片的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPictureFormat : IDisposable
{
    /// <summary>
    /// 获取创建此图片格式设置的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此图片格式设置的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此图片格式设置的父对象。
    /// </summary>
    /// <value>表示此图片格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 按指定增量增加图片的亮度。
    /// </summary>
    /// <param name="increment">亮度增量值（-1.0 到 1.0）。</param>
    void IncrementBrightness(float increment);

    /// <summary>
    /// 按指定增量增加图片的对比度。
    /// </summary>
    /// <param name="increment">对比度增量值（-1.0 到 1.0）。</param>
    void IncrementContrast(float increment);

    /// <summary>
    /// 获取或设置图片的亮度。
    /// </summary>
    /// <value>表示亮度值的浮点数（0.0 到 1.0）。</value>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图片的颜色类型。
    /// </summary>
    /// <value>表示图片颜色类型的 <see cref="MsoPictureColorType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPictureColorType ColorType { get; set; }

    /// <summary>
    /// 获取或设置图片的对比度。
    /// </summary>
    /// <value>表示对比度值的浮点数（0.0 到 1.0）。</value>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置从图片底部裁剪的量（以磅为单位）。
    /// </summary>
    /// <value>表示底部裁剪量的浮点数。</value>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取或设置从图片左侧裁剪的量（以磅为单位）。
    /// </summary>
    /// <value>表示左侧裁剪量的浮点数。</value>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置从图片右侧裁剪的量（以磅为单位）。
    /// </summary>
    /// <value>表示右侧裁剪量的浮点数。</value>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置从图片顶部裁剪的量（以磅为单位）。
    /// </summary>
    /// <value>表示顶部裁剪量的浮点数。</value>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置透明颜色的 RGB 值。
    /// </summary>
    /// <value>表示透明颜色 RGB 值的整数值。</value>
    int TransparencyColor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否将背景设置为透明。
    /// </summary>
    /// <value>指示背景是否透明的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 获取图片的裁剪设置对象。
    /// </summary>
    /// <value>表示裁剪设置的 <see cref="IOfficeCrop"/> 对象。</value>
    IOfficeCrop? Crop { get; }
}