//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中数学对象的分组字符功能接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathGroupChar : IOfficeObject<IWordOMathGroupChar>, IDisposable
{
    /// <summary>
    /// 获取与当前对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置与分组字符关联的数学对象
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置分组字符的字符代码值
    /// </summary>
    short Char { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示分组字符是否显示在顶部
    /// </summary>
    bool CharTop { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示内容是否与顶部对齐
    /// </summary>
    bool AlignTop { get; set; }
}