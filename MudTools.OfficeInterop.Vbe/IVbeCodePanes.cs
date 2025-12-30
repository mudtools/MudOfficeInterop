//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 表示 VBA 编辑器中的代码窗格集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsVb"), ItemIndex]
public interface IVbeCodePanes : IEnumerable<IVbeCodePane?>, IOfficeObject<IVbeCodePanes>, IDisposable
{
    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取此代码窗格集合的父对象（VBE 环境）。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? Parent { get; }

    /// <summary>
    /// 获取集合中代码窗格的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 返回一个可遍历集合的枚举器。
    /// </summary>
    /// <returns>用于遍历集合的枚举器。</returns>
    IVbeCodePane? this[int index] { get; }

    /// <summary>
    /// 返回一个可遍历集合的枚举器。
    /// </summary>
    /// <returns>用于遍历集合的枚举器。</returns>
    IVbeCodePane? this[string name] { get; }

    /// <summary>
    /// 获取或设置当前活动的代码窗格。
    /// </summary>
    IVbeCodePane? Current { get; set; }

}