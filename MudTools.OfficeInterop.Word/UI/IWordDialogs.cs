//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Microsoft Word 中内置对话框的对象集合。
/// <para>注：Dialogs 集合包含 Microsoft Word 中的所有内置对话框。</para>
/// <para>注：使用 Dialogs 属性可返回 Dialogs 集合。</para>
/// <para>注：使用 Dialogs(index)（其中 index 是 WdWordDialog 常量之一）可返回单个 Dialog 对象。</para>
/// </summary>
public interface IWordDialogs : IEnumerable<IWordDialog>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的内置对话框数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 对话框集合属性 (Dialogs Collection Properties)

    /// <summary>
    /// 通过 WdWordDialog 常量获取单个内置对话框。
    /// </summary>
    /// <param name="index">标识对话框的 WdWordDialog 常量。</param>
    /// <returns>指定的 Dialog 对象。</returns>
    IWordDialog this[WdWordDialog index] { get; }

    #endregion
}