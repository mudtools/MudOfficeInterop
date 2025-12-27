//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示所有自动图文集条目的集合。
/// <para>注：使用 AutoCaptions 属性可返回 AutoCaptions 集合。</para>
/// <para>注：使用 AutoCaptions(index)（其中 index 是类名或索引号）可返回单个 AutoCaption 对象。</para>
/// </summary>
public interface IWordAutoCaptions : IEnumerable<IWordAutoCaption>, IDisposable
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
    /// 获取集合中的自动图文集条目数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号或类名获取单个自动图文集条目。
    /// </summary>
    /// <param name="index">索引号（整数）或类名（字符串）。</param>
    /// <returns>指定的自动图文集条目对象。</returns>
    IWordAutoCaption this[object index] { get; }

    #endregion

    #region 自动题注方法 (AutoCaptions Methods)

    /// <summary>
    /// 为指定项目类的所有新插入实例自动插入题注。
    /// </summary>
    /// <param name="name">要为其自动插入题注的项目类名。</param>
    /// <param name="autoInsert">如果为 true，则自动插入题注。</param>
    /// <param name="captionLabel">要使用的题注标签。</param>
    void AutoInsert(string name, bool autoInsert, string captionLabel);

    /// <summary>
    /// 取消选中“插入”菜单上的“自动图文集”命令和“自动套用格式”选项卡上的“自动图文集”复选框。
    /// </summary>
    void CancelAutoInsert();

    #endregion
}
