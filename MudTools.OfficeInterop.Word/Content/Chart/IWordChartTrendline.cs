//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表趋势线的封装接口。
/// </summary>
public interface IWordChartTrendline : IDisposable
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
    /// 获取趋势线名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取趋势线索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置趋势线类型。
    /// </summary>
    XlTrendlineType Type { get; set; }

    /// <summary>
    /// 获取或设置趋势线顺序。
    /// </summary>
    int Order { get; set; }

    /// <summary>
    /// 获取或设置趋势线周期。
    /// </summary>
    int Period { get; set; }

    /// <summary>
    /// 获取或设置是否显示公式。
    /// </summary>
    bool DisplayEquation { get; set; }

    /// <summary>
    /// 获取或设置是否显示 R 平方值。
    /// </summary>
    bool DisplayRSquared { get; set; }

    /// <summary>
    /// 获取或设置是否向前延伸。
    /// </summary>
    double Forward { get; set; }

    /// <summary>
    /// 获取或设置是否向后延伸。
    /// </summary>
    double Backward { get; set; }

    /// <summary>
    /// 获取或设置 Y 轴截距值。
    /// </summary>
    double Intercept { get; set; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取数据标签。
    /// </summary>
    IWordChartDataLabel? DataLabel { get; }

    /// <summary>
    /// 选择趋势线。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除趋势线。
    /// </summary>
    void Delete();
}