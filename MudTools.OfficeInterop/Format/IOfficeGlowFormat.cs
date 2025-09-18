//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中发光格式的接口封装。
/// 该接口提供对形状发光效果的访问和设置。
/// </summary>
public interface IOfficeGlowFormat : IDisposable
{
    /// <summary>
    /// 获取或设置发光效果的颜色格式。
    /// </summary>
    IOfficeColorFormat Color { get; }

    /// <summary>
    /// 获取或设置发光效果的半径大小（单位：磅）。
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// 获取或设置发光效果的透明度（0.0 到 1.0 之间）。
    /// </summary>
    float Transparency { get; set; }
}