//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office命令栏控件集合的接口，提供对命令栏上控件的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCommandBarControls : IOfficeObject<IOfficeCommandBarControls>, IDisposable, IEnumerable<IOfficeCommandBarControl?>
{
    /// <summary>
    /// 获取控件集合中的控件数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取控件（索引从1开始）
    /// </summary>
    /// <param name="index">控件索引</param>
    /// <returns>控件对象</returns>
    IOfficeCommandBarControl? this[int index] { get; }

    /// <summary>
    /// 根据索引获取控件（索引从1开始）
    /// </summary>
    /// <param name="name">控件索引</param>
    /// <returns>控件对象</returns>
    IOfficeCommandBarControl? this[string name] { get; }

    /// <summary>
    /// 添加新的控件到集合中
    /// </summary>
    /// <param name="type">控件类型</param>
    /// <param name="id">控件ID</param>
    /// <param name="parameter">参数</param>
    /// <param name="before">插入位置</param>
    /// <param name="temporary">是否为临时控件</param>
    /// <returns>新创建的控件对象</returns>
    IOfficeCommandBarControl? Add(MsoControlType? type = MsoControlType.msoControlButton,
                                int? id = null, object? parameter = null,
                                object? before = null, bool? temporary = false);
    /// <summary>
    /// 获取父级命令栏
    /// </summary>
    IOfficeCommandBar? Parent { get; }
}