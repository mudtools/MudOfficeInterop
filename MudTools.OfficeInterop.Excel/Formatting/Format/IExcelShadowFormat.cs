//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ShadowFormat 对象的二次封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelShadowFormat : IDisposable
{
    /// <summary>
    /// 获取图片格式对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取与该图片格式相关联的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置阴影类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShadowType Type { get; set; }

    /// <summary>
    /// 获取或设置阴影前景色
    /// </summary>
    IExcelColorFormat? ForeColor { get; set; }

    /// <summary>
    /// 获取或设置阴影大小
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置阴影模糊度
    /// </summary>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置阴影样式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShadowStyle Style { get; set; }

    /// <summary>
    /// 获取或设置阴影是否被遮挡
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Obscured { get; set; }

    /// <summary>
    /// 获取或设置阴影是否随形状旋转
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotateWithShape { get; set; }

    /// <summary>
    /// 获取或设置阴影透明度
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置阴影偏移X坐标
    /// </summary>
    float OffsetX { get; set; }

    /// <summary>
    /// 获取或设置阴影偏移Y坐标
    /// </summary>
    float OffsetY { get; set; }

    /// <summary>
    /// 获取或设置是否可见
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 按指定增量增加阴影X轴偏移量
    /// </summary>
    /// <param name="Increment">增加的偏移量</param>
    void IncrementOffsetX(float Increment);

    /// <summary>
    /// 按指定增量增加阴影Y轴偏移量
    /// </summary>
    /// <param name="Increment">增加的偏移量</param>
    void IncrementOffsetY(float Increment);
}