//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定将选定内容或插入点移动到相对于对象或自身的位置
/// </summary>
[Guid("A1D2A478-67C7-3491-9E7E-64C6E8D43738")]
public enum WdGoToDirection
{
    /// <summary>
    /// 指定对象的第一个实例
    /// </summary>
    wdGoToFirst = 1,

    /// <summary>
    /// 指定对象的最后一个实例
    /// </summary>
    wdGoToLast = -1,

    /// <summary>
    /// 指定对象的下一个实例
    /// </summary>
    wdGoToNext = 2,

    /// <summary>
    /// 相对于当前位置的位置
    /// </summary>
    wdGoToRelative = 2,

    /// <summary>
    /// 指定对象的上一个实例
    /// </summary>
    wdGoToPrevious = 3,

    /// <summary>
    /// 绝对位置
    /// </summary>
    wdGoToAbsolute = 1
}