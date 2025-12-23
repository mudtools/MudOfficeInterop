//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Microsoft.Office.Interop.Word.CoAuthLock 对象的锁定类型
/// </summary>
public enum WdLockType
{
    /// <summary>
    /// 预留供将来使用
    /// </summary>
    wdLockNone,

    /// <summary>
    /// 指定保留锁定。保留锁定是用户通过 Word 2010 中"审阅"选项卡上的"阻止作者"按钮显式创建的。
    /// </summary>
    wdLockReservation,

    /// <summary>
    /// 指定临时锁定。当用户在启用共同创作的文档中开始编辑某个范围时，Word 2010 会自动隐式地对该范围应用临时锁定。
    /// </summary>
    wdLockEphemeral,

    /// <summary>
    /// 指定占位符锁定。占位符锁定表示另一个用户已从该范围移除了他们的锁定，但当前用户尚未通过保存来更新他们对文档的视图。
    /// </summary>
    wdLockChanged
}