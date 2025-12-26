//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中的上标下标数学对象接口，用于处理包含基底、下标和上标的数学公式元素
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathScrSubSup : IOfficeObject<IWordOMathScrSubSup>, IDisposable
{
    /// <summary>
    /// 获取与此数学对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置基底数学表达式元素
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置下标数学表达式元素
    /// </summary>
    IWordOMath? Sub { get; }

    /// <summary>
    /// 获取或设置上标数学表达式元素
    /// </summary>
    IWordOMath? Sup { get; }

    /// <summary>
    /// 获取或设置是否对齐脚本（上标和下标）的值
    /// </summary>
    bool AlignScripts { get; set; }

    /// <summary>
    /// 移除下标元素并返回移除的数学函数对象
    /// </summary>
    /// <returns>移除的下标函数对象，如果不存在则返回null</returns>
    IWordOMathFunction? RemoveSub();

    /// <summary>
    /// 移除上标元素并返回移除的数学函数对象
    /// </summary>
    /// <returns>移除的上标函数对象，如果不存在则返回null</returns>
    IWordOMathFunction? RemoveSup();

    /// <summary>
    /// 将当前的上下标格式转换为前置格式的数学对象
    /// </summary>
    /// <returns>转换后的前置格式数学函数对象，如果转换失败则返回null</returns>
    IWordOMathFunction? ToScrPre();
}