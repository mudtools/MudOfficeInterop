//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中所有邮件合并域的集合的二次封装接口。
/// 此接口允许枚举、访问特定域，并向文档中添加新的邮件合并域 [[1]]。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordMailMergeFields : IEnumerable<IWordMailMergeField?>, IOfficeObject<IWordMailMergeFields>, IDisposable
{
    /// <summary>
    /// 获取此邮件合并域集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此邮件合并域集合的父对象（通常是 <see cref="IWordRange"/> 或 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档中邮件合并域的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的邮件合并域。索引从 1 开始。
    /// </summary>
    /// <param name="index">邮件合并域的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMailMergeField"/> 对象，如果索引无效则返回 null。</returns>
    IWordMailMergeField? this[int index] { get; }

    /// <summary>
    /// 在指定的范围内向文档添加一个新的邮件合并域。
    /// 新域将被插入到范围的起始位置 [[18]]。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="fieldName">要引用的数据源字段名称。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="range"/> 或 <paramref name="fieldName"/> 为 null 或空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加邮件合并域操作失败时抛出。</exception>
    IWordMailMergeField? Add(IWordRange range, string fieldName);

    /// <summary>
    /// 向文档中添加 Ask 邮件合并域。Ask 域会在执行邮件合并时提示用户输入信息。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="name">Ask 域的名称。</param>
    /// <param name="prompt">提示用户输入信息的文本。</param>
    /// <param name="defaultAskText">Ask 域的默认文本。</param>
    /// <param name="askOnce">如果为 true，则仅提示用户一次并在整个合并过程中使用该值。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddAsk(IWordRange range, string name, string? prompt = null, string? defaultAskText = null, bool? askOnce = null);

    /// <summary>
    /// 向文档中添加 Fill-In 邮件合并域。Fill-In 域会提示用户输入信息以填充空白。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="prompt">提示用户输入信息的文本。</param>
    /// <param name="defaultFillInText">Fill-In 域的默认文本。</param>
    /// <param name="askOnce">如果为 true，则仅提示用户一次并在整个合并过程中使用该值。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddFillIn(IWordRange range, string? prompt = null, string? defaultFillInText = null, bool? askOnce = null);

    /// <summary>
    /// 向文档中添加 IF 邮件合并域。IF 域根据条件比较结果决定显示哪个文本。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="mergeField">要比较的邮件合并字段名称。</param>
    /// <param name="comparison">比较操作类型，参考 <see cref="WdMailMergeComparison"/> 枚举。</param>
    /// <param name="compareTo">要与邮件合并字段比较的字符串。</param>
    /// <param name="trueAutoText">当条件为真时使用的自动图文集条目。</param>
    /// <param name="trueText">当条件为真时显示的文本。</param>
    /// <param name="falseAutoText">当条件为假时使用的自动图文集条目。</param>
    /// <param name="falseText">当条件为假时显示的文本。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddIf(IWordRange range, string mergeField, WdMailMergeComparison comparison, string? compareTo = null, string? trueAutoText = null, string? trueText = null, string? falseAutoText = null, string? falseText = null);

    /// <summary>
    /// 向文档中添加 MergeRec 邮件合并域。MergeRec 域会插入数据源记录号。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddMergeRec(IWordRange range);

    /// <summary>
    /// 向文档中添加 MergeSeq 邮件合并域。MergeSeq 域会插入当前记录的序列号。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddMergeSeq(IWordRange range);

    /// <summary>
    /// 向文档中添加 Next 邮件合并域。Next 域会使数据源移动到下一条记录。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddNext(IWordRange range);

    /// <summary>
    /// 向文档中添加 NextIf 邮件合并域。NextIf 域会根据条件比较结果决定是否移到下一条记录。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="MergeField">要比较的邮件合并字段名称。</param>
    /// <param name="Comparison">比较操作类型，参考 <see cref="WdMailMergeComparison"/> 枚举。</param>
    /// <param name="CompareTo">要与邮件合并字段比较的字符串。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddNextIf(IWordRange range, string MergeField, WdMailMergeComparison Comparison, string? CompareTo = null);

    /// <summary>
    /// 向文档中添加 Set 邮件合并域。Set 域会为 Ask 或 Fill-In 域设置默认值。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="Name">要为其设置值的 Ask 或 Fill-In 域的名称。</param>
    /// <param name="valueText">要设置的文本值。</param>
    /// <param name="valueAutoText">要设置的自动图文集条目。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddSet(IWordRange range, string Name, string? valueText = null, string? valueAutoText = null);

    /// <summary>
    /// 向文档中添加 SkipIf 邮件合并域。SkipIf 域会根据条件比较结果决定是否跳过记录。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="mergeField">要比较的邮件合并字段名称。</param>
    /// <param name="comparison">比较操作类型，参考 <see cref="WdMailMergeComparison"/> 枚举。</param>
    /// <param name="compareTo">要与邮件合并字段比较的字符串。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    IWordMailMergeField? AddSkipIf(IWordRange range, string mergeField, WdMailMergeComparison comparison, string? compareTo = null);


}