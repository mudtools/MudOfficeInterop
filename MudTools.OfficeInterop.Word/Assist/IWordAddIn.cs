//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示单个加载宏，无论其当前是否已安装。
/// <para>注：AddIn 对象是 AddIns 集合的成员。</para>
/// <para>注：AddIns 集合包含 Microsoft Word 可用的所有加载项，无论它们当前是否已加载。</para>
/// <para>注：AddIns 集合包括显示在“模板和外接程序”对话框中的全局模板或 Word 外接程序库 (WWL)。</para>
/// </summary>
public interface IWordAddIn : IDisposable
{
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
    /// 获取加载项在 AddIns 集合中的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示启动 Microsoft Word 时是否自动加载指定的加载项。
    /// </summary>
    bool Autoload { get; }

    /// <summary>
    /// 获取一个值，该值指示指定的外接程序是否是 Microsoft Word 外接程序库 (WLL)。
    /// </summary>
    bool Compiled { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示指定的加载项是否已安装（即是否在启动时加载）。
    /// </summary>
    bool Installed { get; set; }

    /// <summary>
    /// 获取指定 AddIn 对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取指定 AddIn 对象的磁盘或 Web 路径。
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 删除指定的 AddIn 对象。
    /// </summary>
    void Delete();
}