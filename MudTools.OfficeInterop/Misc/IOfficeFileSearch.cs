//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中文件搜索功能的接口封装。
/// 该接口提供对文件搜索功能的完整访问。
/// </summary>
public interface IOfficeFileSearch : IDisposable
{
    /// <summary>
    /// 获取搜索结果文件集合。
    /// </summary>
    IOfficeFoundFiles? FoundFiles { get; }

    /// <summary>
    /// 获取搜索范围集合。
    /// </summary>
    IOfficeSearchScopes? SearchScopes { get; }

    /// <summary>
    /// 获取属性测试条件集合。
    /// </summary>
    IOfficePropertyTests? PropertyTests { get; }

    /// <summary>
    /// 获取或设置要搜索的文件类型集合。
    /// </summary>
    IOfficeFileTypes? FileTypes { get; }


    /// <summary>
    /// 获取或设置文件最后修改日期。
    /// </summary>
    MsoLastModified? LastModified { get; set; }

    /// <summary>
    /// 获取或设置文件名。
    /// </summary>
    string FileName { get; set; }

    /// <summary>
    /// 获取搜索文件夹集合。
    /// </summary>
    IOfficeSearchFolders? SearchFolders { get; }

    /// <summary>
    /// 获取或设置是否匹配所有单词形式。
    /// </summary>
    bool MatchAllWordForms { get; set; }


    /// <summary>
    /// 获取或设置是否精确匹配文本。
    /// </summary>
    bool MatchTextExactly { get; set; }

    /// <summary>
    /// 获取或设置是否搜索子文件夹。
    /// </summary>
    bool SearchSubFolders { get; set; }

    /// <summary>
    /// 获取或设置搜索位置。
    /// </summary>
    string LookIn { get; set; }

    /// <summary>
    /// 获取或设置要搜索的文本或属性。
    /// </summary>
    string TextOrProperty { get; set; }

    /// <summary>
    /// 刷新搜索范围。
    /// </summary>
    void RefreshScopes();

    /// <summary>
    /// 获取搜索的文件数量。
    /// </summary>
    int FoundFilesCount { get; }

    /// <summary>
    /// 获取搜索的属性测试条件数量。
    /// </summary>
    int PropertyTestsCount { get; }

    /// <summary>
    /// 执行文件搜索。
    /// </summary>
    /// <param name="sortBy">排序方式。</param>
    /// <param name="alwaysAccurate"></param>
    /// <param name="sortOrder"></param>
    /// <returns>找到的文件数量。</returns>
    int Execute(MsoSortBy sortBy = MsoSortBy.msoSortByFileName,
                       MsoSortOrder sortOrder = MsoSortOrder.msoSortOrderAscending,
                       bool alwaysAccurate = true);

    /// <summary>
    /// 重置搜索条件。
    /// </summary>
    void NewSearch();
}