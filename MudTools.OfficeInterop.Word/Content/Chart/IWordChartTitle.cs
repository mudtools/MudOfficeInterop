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
public interface IWordChartTitle : IDisposable
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
    /// 获取或设置标题文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置标题位置。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    string Caption { get; set; }

    string Formula { get; set; }

    string FormulaLocal { get; set; }

    string FormulaR1C1 { get; set; }

    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取字符对象。
    /// </summary>
    IWordChartCharacters? Characters { get; }

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
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置垂直对齐方式。
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置自动缩放。
    /// </summary>
    bool AutoScaleFont { get; set; }

    bool IncludeInLayout { get; set; }

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
}