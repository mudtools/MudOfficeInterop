//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中带有上限和下限的数学对象接口
/// 此接口用于处理Word文档中的数学公式，特别是带有上下限的数学符号（如积分、求和等）
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathLimUpp : IOfficeObject<IWordOMathLimUpp>, IDisposable
{
    /// <summary>
    /// 获取与此数学对象关联的Word应用程序实例
    /// </summary>
    /// <value>返回IWordApplication接口的实例，如果未关联则返回null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象的父级对象
    /// </summary>
    /// <value>返回父级对象，通常是包含此数学对象的容器</value>
    object? Parent { get; }

    /// <summary>
    /// 获取数学对象的主体部分（通常是积分符号、求和符号等的主要部分）
    /// </summary>
    /// <value>返回IWordOMath接口的实例，代表数学对象的主体部分</value>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取数学对象的上限部分（通常是在积分符号、求和符号上方显示的值）
    /// </summary>
    /// <value>返回IWordOMath接口的实例，代表数学对象的上限部分</value>
    IWordOMath? Lim { get; }

    /// <summary>
    /// 将当前的带有限值的上限数学对象转换为带有限值的下限数学对象
    /// </summary>
    /// <returns>返回IWordOMathFunction接口的实例，代表转换为下限形式的数学函数；如果转换失败则返回null</returns>
    IWordOMathFunction? ToLimLow();
}