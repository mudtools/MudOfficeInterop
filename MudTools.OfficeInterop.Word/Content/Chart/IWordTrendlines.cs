//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表趋势线集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordTrendlines : IEnumerable<IWordTrendline?>, IDisposable
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
    /// 获取趋势线数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取趋势线。
    /// </summary>
    IWordTrendline? this[int index] { get; }


    /// <summary>
    /// 向图表中添加趋势线。
    /// </summary>
    /// <param name="type">趋势线类型，默认为线性趋势线。</param>
    /// <param name="order">趋势线阶数，仅当趋势线类型为多项式时有效。</param>
    /// <param name="period">趋势线周期，仅当趋势线类型为移动平均线时有效。</param>
    /// <param name="forward">趋势线向前延伸的周期数。</param>
    /// <param name="backward">趋势线向后延伸的周期数。</param>
    /// <param name="intercept">趋势线在Y轴上的截距。</param>
    /// <param name="displayEquation">是否在图表上显示趋势线公式。</param>
    /// <param name="displayRSquared">是否在图表上显示R平方值。</param>
    /// <param name="name">趋势线的名称。</param>
    /// <returns>新创建的趋势线对象，如果添加失败则返回null。</returns>
    IWordTrendline? Add(XlTrendlineType type = XlTrendlineType.xlLinear,
                        int? order = null, int? period = null, double? forward = null, double? backward = null,
                        double? intercept = null, bool? displayEquation = null, bool? displayRSquared = null, string? name = null);


}