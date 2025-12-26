//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表轴标题的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordAxisTitle :  IOfficeObject<IWordAxisTitle>,IDisposable
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
    /// 获取或设置轴标题文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取轴标题名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置是否自动缩放字体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置是否在布局中包含轴标题。
    /// </summary>
    bool IncludeInLayout { get; set; }

    /// <summary>
    /// 获取或设置轴标题在图表中的位置。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取或设置轴标题的公式（A1样式引用）。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置轴标题的公式（R1C1样式引用）。
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置轴标题的本地化公式（A1样式引用）。
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置轴标题的本地化公式（R1C1样式引用）。
    /// </summary>
    string FormulaR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置轴标题的阅读顺序。
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
    /// 获取或设置水平对齐方式。
    /// </summary>
    XlConstants HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置垂直对齐方式。
    /// </summary>
    XlConstants VerticalAlignment { get; set; }

    /// <summary>
    /// 获取字符对象。
    /// </summary>
    IWordChartCharacters? Characters { get; }

    /// <summary>
    /// 获取字体格式。
    /// </summary>
    IWordChartFont? Font { get; }

    /// <summary>
    /// 获取内部区域格式。
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取填充格式。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 选择轴标题。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除轴标题。
    /// </summary>
    void Delete();
}