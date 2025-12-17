//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中图片的格式设置接口，提供对图片的各种格式属性和操作方法的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPictureFormat : IDisposable
{
    /// <summary>
    /// 获取图片格式对象的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取与该图片格式相关联的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置图片的亮度值
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图片的颜色类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPictureColorType ColorType { get; set; }

    /// <summary>
    /// 获取或设置图片的对比度值
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置从图片底部裁剪的量
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取或设置从图片左侧裁剪的量
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置从图片右侧裁剪的量
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置从图片顶部裁剪的量
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置图片是否具有透明背景
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 按指定增量增加或减少图片的亮度
    /// </summary>
    /// <param name="increment">亮度的增量值，正数增加亮度，负数减少亮度</param>
    void IncrementBrightness(float increment);

    /// <summary>
    /// 按指定增量增加或减少图片的对比度
    /// </summary>
    /// <param name="increment">对比度的增量值，正数增加对比度，负数减少对比度</param>
    void IncrementContrast(float increment);
}