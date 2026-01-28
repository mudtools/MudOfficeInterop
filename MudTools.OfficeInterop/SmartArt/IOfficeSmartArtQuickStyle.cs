//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中 SmartArt 快速样式的接口封装。
/// 该接口提供对单个 SmartArt 快速样式属性的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeSmartArtQuickStyle : IOfficeObject<IOfficeSmartArtQuickStyle, MsCore.SmartArtQuickStyle>, IDisposable
{
    /// <summary>
    /// 获取 SmartArt 快速样式的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取 SmartArt 快速样式的描述。
    /// </summary>
    string Description { get; }

    /// <summary>
    /// 获取 SmartArt 快速样式的类别名称。
    /// </summary>
    string Category { get; }

    /// <summary>
    /// 获取 SmartArt 快速样式的 ID。
    /// </summary>
    string Id { get; }

    /// <summary>
    /// 获取 SmartArt 快速样式的父对象。
    /// </summary>
    object? Parent { get; }
}