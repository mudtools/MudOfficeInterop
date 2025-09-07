//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示分配给指定项的所有组合键。
/// <para>注：使用 KeysBoundTo 属性可返回 KeysBoundTo 集合。</para>
/// <para>注：使用 KeysBoundTo(index)（其中 index 是索引号）可返回单个 KeyBinding 对象。</para>
/// </summary>
public interface IWordKeysBoundTo : IEnumerable<IWordKeyBinding>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中的组合键数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定项的类别。 返回值是一个 WdKeyCategory 常量。
    /// </summary>
    WdKeyCategory KeyCategory { get; }

    /// <summary>
    /// 获取与指定项相关联的命令、内置函数、宏、字体、自动图文集词条、样式或符号。
    /// </summary>
    string Command { get; }

    /// <summary>
    /// 获取指定项的参数。
    /// </summary>
    string CommandParameter { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号获取单个组合键。
    /// </summary>
    /// <param name="index">索引号（从 1 开始）。</param>
    /// <returns>指定的组合键对象。</returns>
    IWordKeyBinding this[int index] { get; }

    #endregion

    #region KeysBoundTo 方法 (KeysBoundTo Methods)

    /// <summary>
    /// 将新的按键分配添加到当前按键分配方案。
    /// </summary>
    /// <param name="keyCode">按键代码。</param>
    /// <param name="keyCode2">第二个按键代码。</param>
    /// <returns>表示添加的按键分配的对象。</returns>
    IWordKeyBinding Key(int keyCode, object keyCode2);

    #endregion
}
