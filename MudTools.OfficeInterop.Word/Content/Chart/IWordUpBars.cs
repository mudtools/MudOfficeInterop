//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表中上涨柱线的封装接口。
/// 上涨柱线是在股票图表或柱形图表中表示数据点之间价格上涨的部分，通常在开盘价低于收盘价时显示。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordUpBars : IDisposable
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
    /// 获取上涨柱线对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取上涨柱线的边框格式设置。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取上涨柱线的图表格式设置（包括大小、位置和效果等格式）。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取上涨柱线的内部区域格式（背景色、图案等）。
    /// </summary>
    IWordInterior Interior { get; }

    /// <summary>
    /// 获取上涨柱线的填充格式设置。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 选中上涨柱线对象。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? Select();

    /// <summary>
    /// 删除上涨柱线对象。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? Delete();
}