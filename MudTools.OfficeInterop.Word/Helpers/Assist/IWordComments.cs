//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Comment 对象的集合，这些对象表示选择、范围或文档中的批注。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordComments : IEnumerable<IWordComment?>, IOfficeObject<IWordComments>, IDisposable
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
    /// 获取批注数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置其批注显示在批注窗格中的审阅者名称。
    /// </summary>
    string ShowBy { get; set; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="index">指示单个对象序数位置的整数。</param>
    /// <returns>指定索引处的 Comment 对象。</returns>
    IWordComment? this[int index] { get; }

    /// <summary>
    /// 返回表示添加到范围的批注的 Comment 对象。
    /// </summary>
    /// <param name="range">必需。Range 对象。要添加批注的范围。</param>
    /// <param name="text">可选项。批注的文本。</param>
    /// <returns>新创建的 Comment 对象。</returns>
    IWordComment? Add(IWordRange range, string? text = null);
}