//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 启用取消键枚举
/// 用于指定当用户按 Escape 键或 Ctrl+Break 组合键时 Word 的处理方式
/// </summary>
public enum WdEnableCancelKey
{
    /// <summary>
    /// 禁用取消键功能
    /// 用户无法通过 Escape 键或 Ctrl+Break 组合键中断宏执行
    /// </summary>
    wdCancelDisabled,
    
    /// <summary>
    /// 中断模式
    /// 当用户按取消键时，立即中断宏的执行
    /// </summary>
    wdCancelInterrupt
}