//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中下限极限数学对象的接口。此接口提供了对极限下限表达式的访问和操作方法。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathLimLow : IOfficeObject<IWordOMathLimLow>, IDisposable
{
    /// <summary>
    /// 获取与此对象关联的Word应用程序实例。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父级对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置极限表达式中的主体数学公式部分。
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置极限表达式中的下限值。
    /// </summary>
    IWordOMath? Lim { get; }

    /// <summary>
    /// 将当前的下限极限对象转换为上限极限对象。
    /// </summary>
    /// <returns>返回一个新的上限极限数学函数对象，如果操作失败则返回null。</returns>
    IWordOMathFunction? ToLimUpp();
}