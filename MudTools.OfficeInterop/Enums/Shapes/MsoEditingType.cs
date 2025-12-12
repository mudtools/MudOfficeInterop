//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状节点的编辑类型
/// </summary>
public enum MsoEditingType
{
    /// <summary>
    /// 自动编辑类型，由应用程序自动决定节点的控制点位置
    /// </summary>
    msoEditingAuto,
    /// <summary>
    /// 角点编辑类型，节点的进入和离开线段形成尖角
    /// </summary>
    msoEditingCorner,
    /// <summary>
    /// 平滑编辑类型，节点的进入和离开线段形成平滑曲线
    /// </summary>
    msoEditingSmooth,
    /// <summary>
    /// 对称编辑类型，节点两侧的控制点保持对称，形成均匀的曲线
    /// </summary>
    msoEditingSymmetric
}