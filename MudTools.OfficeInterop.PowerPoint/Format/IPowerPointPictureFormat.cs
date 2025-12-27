//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 图片格式接口
/// </summary>
public interface IPowerPointPictureFormat : IDisposable
{

    /// <summary>
    /// 获取或设置亮度
    /// </summary>
    float Brightness { get; set; }

    /// <summary>
    /// 获取或设置对比度
    /// </summary>
    float Contrast { get; set; }

    /// <summary>
    /// 获取或设置是否透明背景
    /// </summary>
    bool TransparentBackground { get; set; }

    /// <summary>
    /// 获取或设置裁剪左边缘
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// 获取或设置裁剪右边缘
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// 获取或设置裁剪上边缘
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// 获取或设置裁剪下边缘
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 裁剪图片
    /// </summary>
    void Crop();

    /// <summary>
    /// 重置图片格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置裁剪区域
    /// </summary>
    /// <param name="left">左裁剪</param>
    /// <param name="right">右裁剪</param>
    /// <param name="top">上裁剪</param>
    /// <param name="bottom">下裁剪</param>
    void SetCrop(float left, float right, float top, float bottom);

    /// <summary>
    /// 应用图片样式
    /// </summary>
    /// <param name="styleIndex">样式索引</param>
    void ApplyStyle(int styleIndex);

    /// <summary>
    /// 获取图片信息
    /// </summary>
    /// <returns>图片信息字符串</returns>
    string GetPictureInfo();
}

