//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中上标数学对象的接口，用于处理数学公式中的上标元素
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathScrSup : IOfficeObject<IWordOMathScrSup>, IDisposable
{
    /// <summary>
    /// 获取与此上标对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此上标对象的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置上标的基础数学表达式（底数）
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置上标部分的数学表达式
    /// </summary>
    IWordOMath? Sup { get; }
}