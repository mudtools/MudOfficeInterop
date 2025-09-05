//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示文件搜索操作。
/// <para>注：使用 Application.FileSearch 属性可返回 FileSearch 对象。</para>
/// <para>注：此接口基于对 Office FileSearch 对象模型的理解实现，因为 Word 特定的 FileSearch SDK 文档有限。</para>
/// </summary>
public interface IWordFileSearch : IDisposable
{
    #region 基本属性 (Basic Properties)   

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    #region 文件搜索属性 (File Search Properties)

    /// <summary>
    /// 获取或设置要搜索的文件夹的路径。
    /// </summary>
    string TextOrProperty { get; set; }

    /// <summary>
    /// 获取或设置要搜索的文件名（可以包含通配符）。
    /// </summary>
    string FileName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在子文件夹中搜索。
    /// </summary>
    bool SearchSubFolders { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示搜索结果中是否包含文件夹。
    /// </summary>
    bool MatchTextExactly { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示搜索结果中是否包含文件夹。
    /// </summary>
    bool MatchAllWordForms { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否区分大小写。
    /// </summary>
    bool MatchCase { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否匹配整个单词。
    /// </summary>
    bool MatchWholeWord { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否使用文本匹配。
    /// </summary>
    bool TextMatch { get; set; }

    /// <summary>
    /// 获取或设置要搜索的文本内容。
    /// </summary>
    string TextToSearch { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否按日期搜索。
    /// </summary>
    bool DateToSearch { get; set; }

    /// <summary>
    /// 获取或设置搜索的开始日期。
    /// </summary>
    DateTime DateFrom { get; set; }

    /// <summary>
    /// 获取或设置搜索的结束日期。
    /// </summary>
    DateTime DateTo { get; set; }

    /// <summary>
    /// 获取或设置文件大小下限。
    /// </summary>
    int SizeMin { get; set; }

    /// <summary>
    /// 获取或设置文件大小上限。
    /// </summary>
    int SizeMax { get; set; }

    /// <summary>
    /// 获取或设置文件类型筛选器。
    /// </summary>
    MsoFileType FileType { get; set; }

    /// <summary>
    /// 获取表示搜索结果的 FoundFiles 集合。
    /// </summary>
    IOfficeFoundFiles FoundFiles { get; }

    /// <summary>
    /// 获取或设置属性筛选器（例如：Author:=Smith）。
    /// </summary>
    string LookIn { get; set; }

    /// <summary>
    /// 获取上次搜索操作返回的文件数。
    /// </summary>
    int LastModified { get; }

    #endregion

    #region 文件搜索方法 (File Search Methods)

    /// <summary>
    /// 执行文件搜索操作。
    /// </summary>
    /// <returns>找到的文件数量。</returns>
    int Execute();

    /// <summary>
    /// 重置文件搜索条件为其默认值。
    /// </summary>
    void Reset();

    /// <summary>
    /// 显示“文件搜索”对话框。
    /// </summary>
    /// <returns>用户在对话框中的操作结果。</returns>
    MsoFeatureInstall ShowSearchDialog();

    #endregion
}