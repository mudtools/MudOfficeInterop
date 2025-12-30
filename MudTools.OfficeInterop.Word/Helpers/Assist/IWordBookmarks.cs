//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档书签集合接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordBookmarks : IDisposable, IOfficeObject<IWordBookmarks>, IEnumerable<IWordBookmark?>
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取书签数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置默认排序方式
    /// </summary>
    WdBookmarkSortBy DefaultSorting { get; set; }

    /// <summary>
    /// 添加书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <param name="range">书签范围</param>
    /// <returns>新添加的书签</returns>
    IWordBookmark? Add(string name, IWordRange range);

    /// <summary>
    /// 判断指定名称的书签是否存在
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    bool? Exists(string name);

    /// <summary>
    /// 根据索引获取书签
    /// </summary>
    IWordBookmark? this[int index] { get; }

    /// <summary>
    /// 根据名称获取书签
    /// </summary>
    IWordBookmark? this[string name] { get; }
}