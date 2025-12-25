//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格列集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordColumns : IEnumerable<IWordColumn?>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取列数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取列。
    /// </summary>
    IWordColumn? this[int index] { get; }

    /// <summary>
    /// 获取第一列。
    /// </summary>
    IWordColumn? First { get; }

    /// <summary>
    /// 获取最后一列。
    /// </summary>
    IWordColumn? Last { get; }

    /// <summary>
    /// 获取列的着色设置。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取列的嵌套级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置列的首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置列的首选宽度类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 添加新的列。
    /// </summary>
    /// <param name="beforeColumn">在指定列前添加。</param>
    /// <returns>新创建的列。</returns>
    IWordColumn? Add(IWordColumn? beforeColumn = null);

    /// <summary>
    /// 删除指定索引的列。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选中所有列。
    /// </summary>
    void Select();

    /// <summary>
    /// 设置所有列的宽度。
    /// </summary>
    /// <param name="columnWidth">列的宽度值。</param>
    /// <param name="rulerStyle">标尺样式，用于确定如何调整列宽。</param>
    void SetWidth(float columnWidth, WdRulerStyle rulerStyle);

    /// <summary>
    /// 自动调整所有列的宽度以适应内容或表格。
    /// </summary>
    void AutoFit();

    /// <summary>
    /// 平均分配所有列的宽度，使它们具有相同的宽度值。
    /// </summary>
    void DistributeWidth();
}