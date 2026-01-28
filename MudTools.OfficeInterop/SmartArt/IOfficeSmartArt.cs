//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// SmartArt 图表的抽象接口，提供对布局、样式、颜色、节点等核心功能的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeSmartArt : IOfficeObject<IOfficeSmartArt, MsCore.SmartArt>, IDisposable
{
    /// <summary>
    /// 获取所有节点的集合（包括嵌套子节点）
    /// </summary>
    IOfficeSmartArtNodes? AllNodes { get; }

    /// <summary>
    /// 获取 SmartArt 图表的直接子节点集合
    /// </summary>
    IOfficeSmartArtNodes? Nodes { get; }

    /// <summary>
    /// 获取或设置 SmartArt 图表的布局样式
    /// </summary>
    IOfficeSmartArtLayout? Layout { get; set; }

    /// <summary>
    /// 获取或设置 SmartArt 图表的快速样式
    /// </summary>
    IOfficeSmartArtQuickStyle? QuickStyle { get; }

    /// <summary>
    /// 获取或设置 SmartArt 图表的颜色样式
    /// </summary>
    IOfficeSmartArtColor? Color { get; set; }

    /// <summary>
    /// 获取或设置是否反转 SmartArt 图表的布局方向
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Reverse { get; set; }

    /// <summary>
    /// 重置 SmartArt 为默认布局和内容
    /// </summary>
    void Reset();
}