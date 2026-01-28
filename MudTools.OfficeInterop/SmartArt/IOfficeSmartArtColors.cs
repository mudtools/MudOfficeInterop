//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中 SmartArt 颜色方案集合的接口封装。
/// 该接口提供对 SmartArt 颜色方案集合的访问和管理。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeSmartArtColors : IOfficeObject<IOfficeSmartArtColors, MsCore.SmartArtColors>, IEnumerable<IOfficeSmartArtColor?>, IDisposable
{
    /// <summary>
    /// 获取 SmartArt 颜色方案集合中项的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取 SmartArt 颜色方案（索引从 1 开始）。
    /// </summary>
    /// <param name="index">SmartArt 颜色方案索引。</param>
    /// <returns>SmartArt 颜色方案对象。</returns>
    IOfficeSmartArtColor? this[int index] { get; }

    /// <summary>
    /// 通过 ID 获取 SmartArt 颜色方案。
    /// </summary>
    /// <param name="id">SmartArt 颜色方案 ID。</param>
    /// <returns>SmartArt 颜色方案对象。</returns>
    IOfficeSmartArtColor? this[string id] { get; }

    /// <summary>
    /// 获取 SmartArt 颜色方案集合的父对象。
    /// </summary>
    object? Parent { get; }


}