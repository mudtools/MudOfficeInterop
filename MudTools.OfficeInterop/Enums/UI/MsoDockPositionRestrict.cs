//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定停靠位置限制的枚举，用于限制自定义任务窗格的停靠行为
/// </summary>
public enum MsoDockPositionRestrict
{
    /// <summary>
    /// 无限制 - 可以停靠在任何位置
    /// </summary>
    msoCTPDockPositionRestrictNone,

    /// <summary>
    /// 无变化 - 保持当前停靠状态，不允许更改
    /// </summary>
    msoCTPDockPositionRestrictNoChange,

    /// <summary>
    /// 无水平移动 - 限制水平方向的停靠调整
    /// </summary>
    msoCTPDockPositionRestrictNoHorizontal,

    /// <summary>
    /// 无垂直移动 - 限制垂直方向的停靠调整
    /// </summary>
    msoCTPDockPositionRestrictNoVertical
}