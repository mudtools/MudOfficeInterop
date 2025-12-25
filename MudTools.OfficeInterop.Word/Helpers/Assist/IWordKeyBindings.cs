//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示当前上下文中的自定义键分配的对象集合。
/// <para>注：使用 KeyBindings 属性可返回 KeyBindings 集合。</para>
/// <para>注：使用 KeyBindings(index)（其中 index 是索引号）可返回单个 KeyBinding 对象。</para>
/// <para>注：使用 Add 方法可将 KeyBinding 对象添加到 KeyBindings 集合中。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordKeyBindings : IEnumerable<IWordKeyBinding?>, IDisposable
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
    /// 获取集合中的自定义键分配数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引号获取单个自定义键分配。
    /// </summary>
    /// <param name="index">索引号（从 1 开始）。</param>
    /// <returns>指定的自定义键分配对象。</returns>
    IWordKeyBinding? this[int index] { get; }

    /// <summary>
    /// 获取一个对象，该对象表示指定键绑定的存储位置。此属性可以返回 Document、Template 或 Application 对象。
    /// </summary>
    object Context { get; }

    /// <summary>
    /// 将新的按键分配添加到当前按键分配方案。
    /// </summary>
    /// <param name="keyCategory">按键分配的类别。</param>
    /// <param name="command">与按键分配相关联的命令、内置函数、宏、字体、自动图文集词条、样式或符号。</param>
    /// <param name="keyCode">按键代码。</param>
    /// <param name="keyCode2">第二个按键代码。</param>
    /// <param name="commandParameter">命令参数。</param>
    /// <returns>表示添加的按键分配的对象。</returns>
    IWordKeyBinding? Add(WdKeyCategory keyCategory, string command, [ConvertInt] WdKey keyCode,
                         WdKey? keyCode2 = null, string? commandParameter = null);

    /// <summary>
    /// 清除所有的自定义按键分配方案，并恢复原有的 Microsoft Word 快捷键分配方案。
    /// </summary>
    void ClearAll();

    /// <summary>
    /// 返回一个 KeyBinding 对象，该对象表示指定的自定义组合键。如果不存在的键组合，则此方法将返回 Nothing。
    /// </summary>
    /// <param name="keyCode">按键代码。</param>
    /// <param name="keyCode2">第二个按键代码。</param>
    /// <returns>表示指定组合键的对象，如果不存在则返回 null。</returns>
    IWordKeyBinding? Key([ConvertInt] WdKey keyCode, WdKey? keyCode2 = null);
}
