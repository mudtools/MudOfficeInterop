//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Microsoft Word 可用的所有加载项的对象集合，无论它们当前是否已加载。
/// <para>注：此集合包括显示在“模板和外接程序”对话框中的全局模板或 Word 外接程序库 (WWL)。</para>
/// <para>注：使用 AddIns 属性可返回 AddIns 集合。</para>
/// <para>注：使用 AddIns(index)（其中 index 是外接程序名称或索引号）可返回单个 AddIn 对象。</para>
/// </summary>
public interface IWordAddIns : IEnumerable<IWordAddIn>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的加载项数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（加载项名称或索引号）获取单个加载项。
    /// </summary>
    /// <param name="index">加载项名称（字符串）或索引号（整数）。</param>
    /// <returns>指定的加载项对象。</returns>
    IWordAddIn this[object index] { get; }

    /// <summary>
    /// 将指定的文件添加到可用加载项列表中。
    /// </summary>
    /// <param name="fileName">要添加的加载项的完整路径和文件名。</param>
    /// <param name="install">如果为 true，则安装加载项；如果为 false，则仅将其添加到列表中但不安装。</param>
    /// <returns>表示添加的加载项的对象。</returns>
    IWordAddIn Add(string fileName, object install);

    /// <summary>
    /// 卸载所有已加载的加载项，并可选择性地将它们从 AddIns 集合中删除。
    /// </summary>
    /// <param name="removeFromList">如果为 true，则从列表中删除加载项；如果为 false，则仅卸载但保留在列表中。</param>
    void Unload(bool removeFromList);
}