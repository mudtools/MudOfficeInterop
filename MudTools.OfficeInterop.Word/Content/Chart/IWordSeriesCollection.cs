//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表系列集合的封装接口。
/// </summary>
public interface IWordSeriesCollection : IEnumerable<IWordChartSeries>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取系列数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取系列。
    /// </summary>
    IWordChartSeries this[int index] { get; }

    /// <summary>
    /// 通过名称获取系列。
    /// </summary>
    IWordChartSeries this[string name] { get; }

    /// <summary>
    /// 添加新的数据系列。
    /// </summary>
    /// <param name="source">数据源范围。</param>
    /// <param name="rowcol">数据排列方式。</param>
    /// <param name="seriesLabels">是否包含系列标签。</param>
    /// <param name="categoryLabels">是否包含分类标签。</param>
    /// <param name="bubbleSizes">气泡大小数据范围。</param>
    /// <returns>新创建的系列。</returns>
    IWordChartSeries Add(object source, MsWord.XlRowCol rowcol, bool seriesLabels, bool categoryLabels, object bubbleSizes);

    /// <summary>
    /// 判断是否存在指定名称的系列。
    /// </summary>
    /// <param name="name">系列名称。</param>
    /// <returns>是否存在。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有系列名称列表。
    /// </summary>
    /// <returns>系列名称列表。</returns>
    List<string> GetNames();
}