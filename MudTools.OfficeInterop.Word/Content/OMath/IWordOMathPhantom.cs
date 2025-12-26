//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中数学对象的虚设(Phantom)元素接口，用于控制数学公式中特定元素的显示属性
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathPhantom : IOfficeObject<IWordOMathPhantom>, IDisposable
{
    /// <summary>
    /// 获取与此数学对象关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置被虚设(Phantom)效果影响的数学表达式元素
    /// </summary>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示虚设(Phantom)元素
    /// </summary>
    bool Show { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将元素的宽度设置为零
    /// </summary>
    bool ZeroWid { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将元素的上部间距(Ascender)设置为零
    /// </summary>
    bool ZeroAsc { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将元素的下部间距(Descender)设置为零
    /// </summary>
    bool ZeroDesc { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将元素设置为透明
    /// </summary>
    bool Transp { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否压缩元素之间的间距
    /// </summary>
    bool Smash { get; set; }
}