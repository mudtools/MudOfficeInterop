//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Microsoft Office Word 如何处理出现在换行符前的减法运算符
/// </summary>
public enum WdOMathBreakSub
{
    /// <summary>
    /// 在换行符前以减号结束，并在下一行开头重复该减号。默认值。
    /// </summary>
    wdOMathBreakSubMinusMinus,

    /// <summary>
    /// 在第一行的换行符前插入加号，在下一行的数字前插入减号
    /// </summary>
    wdOMathBreakSubPlusMinus,

    /// <summary>
    /// 在第一行的换行符前插入减号，在下一行的数字前插入加号
    /// </summary>
    wdOMathBreakSubMinusPlus
}