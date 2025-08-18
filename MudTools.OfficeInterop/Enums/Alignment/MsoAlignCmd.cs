//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定对齐命令的枚举，用于对齐形状或对象
/// </summary>
public enum MsoAlignCmd
{
    /// <summary>
    /// 左对齐 - 将所选对象的左边缘对齐
    /// </summary>
    msoAlignLefts,

    /// <summary>
    /// 居中对齐 - 将所选对象的中心对齐
    /// </summary>
    msoAlignCenters,

    /// <summary>
    /// 右对齐 - 将所选对象的右边缘对齐
    /// </summary>
    msoAlignRights,

    /// <summary>
    /// 顶对齐 - 将所选对象的顶部对齐
    /// </summary>
    msoAlignTops,

    /// <summary>
    /// 垂直居中对齐 - 将所选对象的垂直中心对齐
    /// </summary>
    msoAlignMiddles,

    /// <summary>
    /// 底对齐 - 将所选对象的底部对齐
    /// </summary>
    msoAlignBottoms
}