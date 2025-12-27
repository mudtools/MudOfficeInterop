//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 页眉页脚接口
/// </summary>
public interface IPowerPointHeadersFooters : IDisposable
{
    /// <summary>
    /// 获取幻灯片编号
    /// </summary>
    IPowerPointHeaderFooter SlideNumber { get; }

    /// <summary>
    /// 获取日期和时间
    /// </summary>
    IPowerPointHeaderFooter DateAndTime { get; }

    /// <summary>
    /// 获取页脚
    /// </summary>
    IPowerPointHeaderFooter Footer { get; }

    /// <summary>
    /// 获取页眉
    /// </summary>
    IPowerPointHeaderFooter Header { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置是否显示幻灯片编号
    /// </summary>
    bool SlideNumberVisible { get; set; }

    /// <summary>
    /// 获取或设置是否显示日期和时间
    /// </summary>
    bool DateAndTimeVisible { get; set; }

    /// <summary>
    /// 获取或设置是否显示页脚
    /// </summary>
    bool FooterVisible { get; set; }

    /// <summary>
    /// 获取或设置是否显示页眉
    /// </summary>
    bool HeaderVisible { get; set; }

    /// <summary>
    /// 获取或设置日期和时间格式
    /// </summary>
    int DateTimeFormat { get; set; }

    /// <summary>
    /// 获取或设置是否使用预设格式
    /// </summary>
    bool UseDateTimeFormat { get; set; }

    /// <summary>
    /// 获取或设置是否显示背景图形
    /// </summary>
    bool BackgroundVisible { get; set; }

    /// <summary>
    /// 设置日期和时间文本
    /// </summary>
    /// <param name="dateTimeText">日期时间文本</param>
    void SetDateTimeText(string dateTimeText);

    /// <summary>
    /// 设置页脚文本
    /// </summary>
    /// <param name="footerText">页脚文本</param>
    void SetFooterText(string footerText);

    /// <summary>
    /// 设置页眉文本
    /// </summary>
    /// <param name="headerText">页眉文本</param>
    void SetHeaderText(string headerText);

    /// <summary>
    /// 设置幻灯片编号格式
    /// </summary>
    /// <param name="format">编号格式</param>
    void SetSlideNumberFormat(int format);

    /// <summary>
    /// 获取页眉页脚信息
    /// </summary>
    /// <returns>页眉页脚信息字符串</returns>
    string GetHeadersFootersInfo();
}