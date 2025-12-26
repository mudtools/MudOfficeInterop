//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中的数学函数对象，提供对Word数学函数及其相关组件的访问和操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathFunction : IOfficeObject<IWordOMathFunction>, IDisposable
{
    /// <summary>
    /// 获取与该数学函数关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该数学函数的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示该数学函数在文档中的范围对象
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取数学函数的参数集合对象
    /// </summary>
    IWordOMathArgs? Args { get; }

    /// <summary>
    /// 获取数学函数的叠加符号对象（如上划线、下划线等）
    /// </summary>
    IWordOMathAcc? Acc { get; }

    /// <summary>
    /// 获取数学函数的上划线/下划线对象
    /// </summary>
    IWordOMathBar? Bar { get; }

    /// <summary>
    /// 获取数学函数的方框对象
    /// </summary>
    IWordOMathBox? Box { get; }

    /// <summary>
    /// 获取数学函数的边框方框对象
    /// </summary>
    IWordOMathBorderBox? BorderBox { get; }

    /// <summary>
    /// 获取数学函数的分隔符对象（如括号、花括号等）
    /// </summary>
    IWordOMathDelim? Delim { get; }

    /// <summary>
    /// 获取数学函数的等式数组对象
    /// </summary>
    IWordOMathEqArray? EqArray { get; }

    /// <summary>
    /// 获取数学函数的分数对象
    /// </summary>
    IWordOMathFrac? Frac { get; }

    /// <summary>
    /// 获取数学函数的函数对象（表示一个完整的函数结构）
    /// </summary>
    IWordOMathFunc? Func { get; }

    /// <summary>
    /// 获取数学函数的组合字符对象
    /// </summary>
    IWordOMathGroupChar? GroupChar { get; }

    /// <summary>
    /// 获取数学函数的下极限对象
    /// </summary>
    IWordOMathLimLow? LimLow { get; }

    /// <summary>
    /// 获取数学函数的上极限对象
    /// </summary>
    IWordOMathLimUpp? LimUpp { get; }

    /// <summary>
    /// 获取数学函数的矩阵对象
    /// </summary>
    IWordOMathMat? Mat { get; }

    /// <summary>
    /// 获取数学函数的n元运算符对象（如积分、求和等）
    /// </summary>
    IWordOMathNary? Nary { get; }

    /// <summary>
    /// 获取数学函数的幻影对象（用于占位符等特殊用途）
    /// </summary>
    IWordOMathPhantom? Phantom { get; }

    /// <summary>
    /// 获取数学函数的前缀脚本对象
    /// </summary>
    IWordOMathScrPre? ScrPre { get; }

    /// <summary>
    /// 获取数学函数的根式对象
    /// </summary>
    IWordOMathRad? Rad { get; }

    /// <summary>
    /// 获取数学函数的下标对象
    /// </summary>
    IWordOMathScrSub? ScrSub { get; }

    /// <summary>
    /// 获取数学函数的上下标对象（同时包含下标和上标）
    /// </summary>
    IWordOMathScrSubSup? ScrSubSup { get; }

    /// <summary>
    /// 获取数学函数的上标对象
    /// </summary>
    IWordOMathScrSup? ScrSup { get; }

    /// <summary>
    /// 获取数学函数的数学对象
    /// </summary>
    IWordOMath? OMath { get; }

    /// <summary>
    /// 从文档中移除当前数学函数，并返回被移除的对象
    /// </summary>
    /// <returns>被移除的数学函数对象，如果操作失败则返回null</returns>
    IWordOMathFunction? Remove();
}