//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中的上标/下标预设数学对象接口
/// 此接口用于处理包含基元素、上标和下标的数学表达式
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathScrPre : IDisposable
{
    /// <summary>
    /// 获取与此数学对象关联的 Word 应用程序实例
    /// </summary>
    /// <value>返回 IWordApplication 类型的对象，可能为 null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象的父级对象
    /// </summary>
    /// <value>返回父级对象，可能为 null</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置此数学对象的下标部分
    /// </summary>
    /// <value>返回 IWordOMath 类型的下标对象，可能为 null</value>
    IWordOMath? Sub { get; }

    /// <summary>
    /// 获取或设置此数学对象的上标部分
    /// </summary>
    /// <value>返回 IWordOMath 类型的上标对象，可能为 null</value>
    IWordOMath? Sup { get; }

    /// <summary>
    /// 获取或设置此数学对象的基元素（主体部分）
    /// </summary>
    /// <value>返回 IWordOMath 类型的基元素对象，可能为 null</value>
    IWordOMath? E { get; }

    /// <summary>
    /// 将当前的上标/下标对象转换为下标-上标函数格式
    /// </summary>
    /// <returns>返回 IWordOMathFunction 类型的转换结果，可能为 null</returns>
    IWordOMathFunction? ToScrSubSup();
}