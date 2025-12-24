//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示制表符停止位集合的接口，提供对制表符停止位的访问和操作功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeTabStops2 : IEnumerable<IOfficeTabStop2?>, IDisposable
{
    /// <summary>
    /// 获取制表符停止位集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取制表符停止位的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置默认制表符间距
    /// </summary>
    float DefaultSpacing { get; set; }

    /// <summary>
    /// 根据索引获取制表符停止位
    /// </summary>
    /// <param name="index">制表符停止位的索引（从1开始）</param>
    /// <returns>指定索引处的制表符停止位对象</returns>
    IOfficeTabStop2? this[int index] { get; }

    /// <summary>
    /// 添加一个新的制表符停止位
    /// </summary>
    /// <param name="type">制表符类型，参考 <see cref="MsoTabStopType"/> 枚举</param>
    /// <param name="position">制表符的位置（以磅为单位）</param>
    /// <returns>新创建的制表符停止位对象</returns>
    IOfficeTabStop2? Add(MsoTabStopType type, float position);
}