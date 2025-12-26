//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office应用程序中的命令栏集合接口，提供对命令栏集合的枚举和资源管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCommandBars : IEnumerable<IOfficeCommandBar?>, IOfficeObject<IOfficeCommandBars>, IDisposable
{
    /// <summary>
    /// 获取命令栏集合中的命令栏数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取命令栏（索引从1开始）
    /// </summary>
    /// <param name="index">命令栏索引</param>
    /// <returns>命令栏对象</returns>
    IOfficeCommandBar? this[int index] { get; }

    /// <summary>
    /// 根据名称获取命令栏
    /// </summary>
    /// <param name="name">命令栏名称</param>
    /// <returns>命令栏对象</returns>
    IOfficeCommandBar? this[string name] { get; }

    /// <summary>
    /// 添加新的命令栏
    /// </summary>
    /// <param name="name">命令栏名称</param>
    /// <param name="position">命令栏位置</param>
    /// <param name="menuBar">是否为菜单栏</param>
    /// <param name="temporary">是否为临时命令栏</param>
    /// <returns>新创建的命令栏对象</returns>
    IOfficeCommandBar? Add(string? name = null, MsoBarPosition position = MsoBarPosition.msoBarTop,
                         bool? menuBar = false, bool? temporary = false);


    /// <summary>
    /// 查找符合指定条件的单个命令栏控件
    /// </summary>
    /// <param name="type">控件类型</param>
    /// <param name="id">控件ID</param>
    /// <param name="tag">控件标签</param>
    /// <param name="visible">是否可见</param>
    /// <returns>找到的命令栏控件，未找到则返回null</returns>
    IOfficeCommandBarControl? FindControl(MsoControlType? type, string? id, object? tag, bool? visible);


    /// <summary>
    /// 查找符合指定条件的所有命令栏控件
    /// </summary>
    /// <param name="type">控件类型</param>
    /// <param name="id">控件ID</param>
    /// <param name="tag">控件标签</param>
    /// <param name="visible">是否可见</param>
    /// <returns>找到的命令栏控件集合，未找到则返回null</returns>
    IOfficeCommandBarControls? FindControls(MsoControlType? type, string? id, object? tag, bool? visible);

    /// <summary>
    /// 获取或设置大型按钮显示模式
    /// </summary>
    bool LargeButtons { get; set; }

    /// <summary>
    /// 获取或设置菜单动画效果
    /// </summary>
    MsoMenuAnimation MenuAnimationStyle { get; set; }
}