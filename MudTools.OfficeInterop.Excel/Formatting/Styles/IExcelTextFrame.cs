//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel TextFrame 对象的二次封装接口
/// </summary>
public interface IExcelTextFrame : IDisposable
{
    /// <summary>
    /// 获取或设置文本方向
    /// </summary>
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置是否自动调整边距
    /// </summary>
    bool AutoMargins { get; set; }

    /// <summary>
    /// 获取或设置文本阅读顺序
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置自动调整大小
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取或设置水平对齐方式
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置垂直对齐方式
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置文本垂直溢出行为
    /// </summary>
    XlOartVerticalOverflow VerticalOverflow { get; set; }

    /// <summary>
    /// 获取或设置文本水平溢出行为
    /// </summary>
    XlOartHorizontalOverflow HorizontalOverflow { get; set; }

    /// <summary>
    /// 获取或设置边距
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置边距
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置边距
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置边距
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取文本框中指定位置和长度的字符对象
    /// </summary>
    /// <param name="start">起始字符位置（从1开始），默认为null表示从第一个字符开始</param>
    /// <param name="length">要获取的字符数量，默认为null表示获取到文本末尾</param>
    /// <returns>表示指定文本范围的字符对象，如果文本框为空则返回null</returns>
    IExcelCharacters? Characters(int? start = null, int? length = null);
}