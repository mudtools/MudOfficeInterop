//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示一个自定义组合键。
/// <para>注：KeyBinding 对象是 KeyBindings 集合的成员。使用 KeyBindings 集合的 Add 方法可将 KeyBinding 对象添加到 KeyBindings 集合中。</para>
/// <para>注：使用 KeyBindings(index)（其中 index 是索引号）可返回单个 KeyBinding 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordKeyBinding : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }


    /// <summary>
    /// 获取一个值，该值指示此组合键是否为受保护的组合键。
    /// </summary>
    bool Protected { get; }

    /// <summary>
    /// 获取或设置指定的组合键。 返回值是一个 WdKey 常量。
    /// </summary>
    WdKey KeyCode { get; }

    /// <summary>
    /// 获取或设置组合键的第二个键代码。 返回值是一个 WdKey 常量。
    /// </summary>
    WdKey KeyCode2 { get; }

    /// <summary>
    /// 获取指定的组合键。 返回值是一个 WdKey 常量。
    /// </summary>
    WdKeyCategory KeyCategory { get; }

    /// <summary>
    /// 获取或设置与指定的组合键相关联的命令、内置函数、宏、字体、自动图文集词条、样式或符号。
    /// </summary>
    string Command { get; }

    /// <summary>
    /// 获取或设置指定组合键的参数。
    /// </summary>
    string CommandParameter { get; }

    /// <summary>
    /// 获取一个对象，该对象表示指定键绑定的存储位置。此属性可以返回 Document、Template 或 Application 对象。
    /// </summary>
    object Context { get; }

    /// <summary>
    /// 删除指定的自定义组合键。
    /// </summary>
    void Clear();

    /// <summary>
    /// 重置指定的组合键，使其与默认设置相匹配。
    /// </summary>
    void Execute();
}