//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示可用题注标签的 CaptionLabel 对象集合。
/// 此接口提供对题注标签集合的管理功能，包括添加、检索和枚举题注标签。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordCaptionLabels : IEnumerable<IWordCaptionLabel?>, IOfficeObject<IWordCaptionLabels, MsWord.CaptionLabels>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的题注标签数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定的题注标签。
    /// </summary>
    /// <param name="index">指示序号位置或表示单个对象名称的字符串。</param>
    /// <returns>指定索引处的题注标签对象。</returns>
    IWordCaptionLabel? this[int index] { get; }

    /// <summary>
    /// 通过索引获取指定的题注标签。
    /// </summary>
    /// <param name="name">指示序号位置或表示单个对象名称的字符串。</param>
    /// <returns>指定索引处的题注标签对象。</returns>
    IWordCaptionLabel? this[string name] { get; }

    /// <summary>
    /// 将一个表示自定义题注标签的 CaptionLabel 对象添加到 CaptionLabels 集合中。
    /// </summary>
    /// <param name="name">CaptionLabel 对象的名称。</param>
    /// <returns>新创建的自定义题注标签对象。</returns>
    IWordCaptionLabel? Add(string name);
}