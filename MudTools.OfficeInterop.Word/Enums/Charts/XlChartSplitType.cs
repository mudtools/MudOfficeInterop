//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在复合饼图或条饼图的第二个图表中显示的值
/// </summary>
public enum XlChartSplitType
{
    /// <summary>
    /// 第二个图表显示数据系列中的最小值，显示的值数量由 SplitValue 属性指定
    /// </summary>
    xlSplitByPosition = 1,

    /// <summary>
    /// 第二个图表显示小于总值某个百分比的值，百分比由 SplitValue 属性指定
    /// </summary>
    xlSplitByPercentValue = 3,

    /// <summary>
    /// 在第二个图表中显示任意数据切片
    /// </summary>
    xlSplitByCustomSplit = 4,

    /// <summary>
    /// 第二个图表显示小于 SplitValue 属性指定值的值
    /// </summary>
    xlSplitByValue = 2
}