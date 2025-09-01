//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Crop 的接口，用于操作图片裁剪格式。
/// </summary>
public interface IWordCrop : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }


    /// <summary>
    /// 获取或设置形状的左边缘位置（相对于图片容器的左边缘）
    /// </summary>
    float ShapeLeft { get; set; }

    /// <summary>
    /// 获取或设置形状的上边缘位置（相对于图片容器的上边缘）
    /// </summary>
    float ShapeTop { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度
    /// </summary>
    float ShapeWidth { get; set; }

    /// <summary>
    /// 获取或设置形状的高度
    /// </summary>
    float ShapeHeight { get; set; }

    /// <summary>
    /// 获取或设置图片相对于形状的水平偏移量
    /// </summary>
    float PictureOffsetX { get; set; }

    /// <summary>
    /// 获取或设置图片相对于形状的垂直偏移量
    /// </summary>
    float PictureOffsetY { get; set; }

    /// <summary>
    /// 获取或设置图片的宽度
    /// </summary>
    float PictureWidth { get; set; }

    /// <summary>
    /// 获取或设置图片的高度
    /// </summary>
    float PictureHeight { get; set; }

    /// <summary>
    /// 设置形状的尺寸
    /// </summary>
    /// <param name="width">形状的宽度</param>
    /// <param name="height">形状的高度</param>
    void SetShapeSize(float width, float height);
}