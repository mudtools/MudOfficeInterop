//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表标题的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartTitle : IOfficeObject<IWordChartTitle>, IDisposable
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
    /// 获取或设置标题文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置标题位置。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取或设置图表标题的说明文字。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置图表标题的公式表达式。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置图表标题的本地化公式表达式。
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置图表标题的R1C1引用样式公式表达式。
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取图表标题的边框格式设置。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取字体格式。
    /// </summary>
    IWordChartFont? Font { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置水平对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlConstants HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置垂直对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlConstants VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置文本方向。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置自动缩放。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置是否在布局中包含标题。
    /// </summary>
    bool IncludeInLayout { get; set; }

    /// <summary>
    /// 获取或设置阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置左边距。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置顶边距。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置宽度。
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取或设置高度。
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 选择标题对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除标题。
    /// </summary>
    void Delete();

    /// <summary>
    /// 返回一个代表图表标题中文本字符的对象，可用于对部分文本进行格式设置。
    /// </summary>
    /// <param name="start">可选参数，指定要返回的第一个字符的位置（从1开始计数）。如果省略此参数，则包含所有字符。</param>
    /// <param name="length">可选参数，指定要返回的字符数量。如果省略此参数，则包含从起始位置到文本末尾的所有字符。</param>
    /// <returns>返回表示图表标题中指定字符范围的IWordChartCharacters对象。</returns>
    [IgnoreGenerator]
    IWordChartCharacters? Characters(int? start = null, int? length = null);
}