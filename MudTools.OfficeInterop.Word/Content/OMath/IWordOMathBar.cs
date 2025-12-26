//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中数学对象的上划线(bar)元素的接口
/// 在数学公式中，bar通常用于表示上划线符号，例如在复数的共轭或变量的平均值表示中
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathBar : IOfficeObject<IWordOMathBar>, IDisposable
{
    /// <summary>
    /// 获取与当前数学对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前数学对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置构成bar符号下方的数学表达式元素
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置一个布尔值，指示bar符号是否位于上方(true)或下方(false)
    /// </summary>
    bool BarTop { get; set; }
}