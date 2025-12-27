//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 阴影格式接口
/// </summary>
public interface IPowerPointShadowFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置阴影类型
    /// </summary>
    int Type { get; set; }

    /// <summary>
    /// 获取或设置水平偏移
    /// </summary>
    float OffsetX { get; set; }

    /// <summary>
    /// 获取或设置垂直偏移
    /// </summary>
    float OffsetY { get; set; }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    int ForeColor { get; set; }


    /// <summary>
    /// 获取或设置模糊度
    /// </summary>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置阴影大小
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 重置阴影格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制阴影格式
    /// </summary>
    /// <returns>复制的阴影格式对象</returns>
    IPowerPointShadowFormat Duplicate();

    /// <summary>
    /// 应用阴影格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    void ApplyTo(IPowerPointShape shape);

    /// <summary>
    /// 设置阴影偏移
    /// </summary>
    /// <param name="offsetX">水平偏移</param>
    /// <param name="offsetY">垂直偏移</param>
    void SetOffset(float offsetX, float offsetY);

    /// <summary>
    /// 设置阴影颜色
    /// </summary>
    /// <param name="color">阴影颜色</param>
    /// <param name="transparency">透明度</param>
    void SetColor(int color, float transparency = 0);

    /// <summary>
    /// 设置阴影大小和模糊度
    /// </summary>
    /// <param name="size">阴影大小</param>
    /// <param name="blur">模糊度</param>
    void SetSizeAndBlur(float size, float blur);

    /// <summary>
    /// 获取阴影信息
    /// </summary>
    /// <returns>阴影信息字符串</returns>
    string GetShadowInfo();
}
