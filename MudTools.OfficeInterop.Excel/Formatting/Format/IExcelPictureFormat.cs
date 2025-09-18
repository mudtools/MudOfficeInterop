namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中图片的格式设置接口，提供对图片的各种格式属性和操作方法的访问
/// </summary>
public interface IExcelPictureFormat : IDisposable
{
    /// <summary>
    /// 获取图片格式对象的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取与该图片格式相关联的Excel应用程序对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置图片的亮度值
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置图片的颜色类型
    /// </summary>
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
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 按指定增量增加或减少图片的亮度
    /// </summary>
    /// <param name="Increment">亮度的增量值，正数增加亮度，负数减少亮度</param>
    void IncrementBrightness(float Increment);

    /// <summary>
    /// 按指定增量增加或减少图片的对比度
    /// </summary>
    /// <param name="Increment">对比度的增量值，正数增加对比度，负数减少对比度</param>
    void IncrementContrast(float Increment);
}