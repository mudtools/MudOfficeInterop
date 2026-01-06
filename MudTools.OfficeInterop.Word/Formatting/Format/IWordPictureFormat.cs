//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureFormat 的接口，用于操作图片格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordPictureFormat : IOfficeObject<IWordPictureFormat, MsWord.PictureFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置指定图片或 OLE 对象的亮度。此属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数字。
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置应用于指定图片或 OLE 对象的颜色变换类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPictureColorType ColorType { get; set; }

    /// <summary>
    /// 获取或设置指定图片或 OLE 对象的对比度。此属性的值必须是 0.0（最小对比度）到 1.0（最大对比度）之间的数字。
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置从指定图片或 OLE 对象底部裁剪的点数。
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取或设置从指定图片或 OLE 对象左侧裁剪的点数。
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置从指定图片或 OLE 对象右侧裁剪的点数。
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置从指定图片或 OLE 对象顶部裁剪的点数。
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置指定图片的透明颜色作为 RGB 值。仅适用于位图。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color TransparencyColor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示使用透明颜色定义的图片部分是否实际显示为透明。仅适用于位图。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 按指定增量更改图片的亮度。使用 Brightness 属性设置图片的绝对亮度。
    /// </summary>
    /// <param name="increment">必需。指定要更改图片亮度属性值的量。正值使图片更亮；负值使图片更暗。</param>
    void IncrementBrightness(float increment);

    /// <summary>
    /// 按指定增量更改图片的对比度。使用 Contrast 属性设置图片的绝对对比度。
    /// </summary>
    /// <param name="increment">必需。指定要更改图片对比度属性值的量。正值增加对比度；负值减少对比度。</param>
    void IncrementContrast(float increment);

    /// <summary>
    /// 获取或设置表示图像裁剪的 Crop 对象。
    /// </summary>
    IOfficeCrop? Crop { get; set; }
}