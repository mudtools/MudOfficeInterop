//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中数学函数对象的接口，用于操作和访问Word文档中的数学函数元素
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathFunc : IDisposable
{
    /// <summary>
    /// 获取与此数学函数关联的Word应用程序实例
    /// </summary>
    /// <value>返回IWordApplication接口的实例，如果不存在则返回null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学函数的父对象
    /// </summary>
    /// <value>返回父对象，通常是指包含此数学函数的对象</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置函数名称部分的数学对象
    /// </summary>
    /// <value>返回IWordOMath接口的实例，表示函数名称部分，如果不存在则返回null</value>
    IWordOMath? FName { get; }

    /// <summary>
    /// 获取或设置函数参数部分的数学对象
    /// </summary>
    /// <value>返回IWordOMath接口的实例，表示函数参数部分，如果不存在则返回null</value>
    IWordOMath? E { get; }
}