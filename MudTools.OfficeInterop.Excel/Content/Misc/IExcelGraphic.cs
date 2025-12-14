//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示Excel中的图形对象接口，提供对Excel图形元素的基本操作功能。
/// 该接口继承自IDisposable接口，支持资源释放功能，并通过ComObjectWrap特性与Excel COM对象进行交互。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelGraphic : IDisposable
{
    /// <summary>
    /// 获取图形对象的父级对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置图形的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置图形的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置是否锁定图形的纵横比。
    /// 当设置为true时，调整图形大小时会保持原始宽高比；当设置为false时，可以自由调整图形的宽度和高度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取或设置图形的亮度（-1到1）
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图形的对比度（0-1）
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置图形的颜色类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPictureColorType ColorType { get; set; }

    /// <summary>
    /// 获取图形的文件名
    /// </summary>
    string Filename { get; set; }

    /// <summary>
    /// 获取图形的裁剪左边界
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取图形的裁剪右边界
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取图形的裁剪上边界
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取图形的裁剪下边界
    /// </summary>
    float CropBottom { get; set; }

}
