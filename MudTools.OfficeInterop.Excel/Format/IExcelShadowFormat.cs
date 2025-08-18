//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

// <summary>
/// Excel ShadowFormat 对象的二次封装接口
/// </summary>
public interface IExcelShadowFormat : IDisposable
{
    /// <summary>
    /// 获取或设置阴影类型
    /// </summary>
    MsoShadowType Type { get; set; }

    /// <summary>
    /// 获取或设置阴影颜色
    /// </summary>
    int Color { get; set; }

    /// <summary>
    /// 获取或设置阴影透明度
    /// </summary>
    int Transparency { get; set; }

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
    bool Visible { get; set; }
}