//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 框架集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordFrames : IEnumerable<IWordFrame?>, IOfficeObject<IWordFrames, MsWord.Frames>, IDisposable
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
    /// 获取框架数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取框架。
    /// </summary>
    IWordFrame? this[int index] { get; }

    /// <summary>
    /// 添加新的框架。
    /// </summary>
    /// <param name="range">框架范围。</param>
    /// <returns>新创建的框架。</returns>
    IWordFrame? Add(IWordRange range);

    /// <summary>
    /// 删除指定索引的框架。
    /// </summary>
    void Delete();

}