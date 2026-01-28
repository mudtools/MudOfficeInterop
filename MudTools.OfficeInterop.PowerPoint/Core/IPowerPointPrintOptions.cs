//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 打印选项接口
/// </summary>
public interface IPowerPointPrintOptions : IDisposable
{
    /// <summary>
    /// 获取或设置打印份数
    /// </summary>
    int Copies { get; set; }

    /// <summary>
    /// 获取或设置是否逐份打印
    /// </summary>
    bool Collate { get; set; }

    /// <summary>
    /// 获取或设置打印范围
    /// </summary>
    string PrintRange { get; set; }

    /// <summary>
    /// 获取或设置打印到文件
    /// </summary>
    bool PrintToFile { get; set; }

    /// <summary>
    /// 获取或设置输出文件名
    /// </summary>
    string OutputFileName { get; set; }

    /// <summary>
    /// 获取或设置打印机名称
    /// </summary>
    string PrinterName { get; set; }

    /// <summary>
    /// 获取或设置打印质量
    /// </summary>
    int PrintQuality { get; set; }

    /// <summary>
    /// 获取或设置打印颜色模式
    /// </summary>
    int ColorMode { get; set; }

    /// <summary>
    /// 获取或设置打印方向
    /// </summary>
    int Orientation { get; set; }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    int PaperSize { get; set; }

    /// <summary>
    /// 获取或设置纸张来源
    /// </summary>
    int PaperSource { get; set; }

    /// <summary>
    /// 获取或设置打印分辨率
    /// </summary>
    int Resolution { get; set; }

    /// <summary>
    /// 获取或设置打印份数
    /// </summary>
    int NumberOfCopies { get; set; }

    /// <summary>
    /// 获取或设置是否双面打印
    /// </summary>
    bool Duplex { get; set; }

    /// <summary>
    /// 获取或设置是否打印注释
    /// </summary>
    bool PrintComments { get; set; }

    /// <summary>
    /// 获取或设置是否打印隐藏幻灯片
    /// </summary>
    bool PrintHiddenSlides { get; set; }

    /// <summary>
    /// 获取或设置是否打印幻灯片编号
    /// </summary>
    bool PrintSlideNumbers { get; set; }

    /// <summary>
    /// 获取或设置是否打印日期和时间
    /// </summary>
    bool PrintDateTime { get; set; }

    /// <summary>
    /// 获取或设置是否打印页眉页脚
    /// </summary>
    bool PrintHeadersFooters { get; set; }

    /// <summary>
    /// 获取或设置是否打印备注
    /// </summary>
    bool PrintNotes { get; set; }

    /// <summary>
    /// 获取或设置是否打印讲义
    /// </summary>
    bool PrintHandouts { get; set; }

    /// <summary>
    /// 获取或设置是否打印大纲
    /// </summary>
    bool PrintOutline { get; set; }

    /// <summary>
    /// 获取或设置是否打印幻灯片
    /// </summary>
    bool PrintSlides { get; set; }

    /// <summary>
    /// 获取或设置是否打印备注页
    /// </summary>
    bool PrintNotesPages { get; set; }

    /// <summary>
    /// 获取或设置是否打印讲义页
    /// </summary>
    bool PrintHandoutPages { get; set; }

    /// <summary>
    /// 获取或设置是否打印大纲页
    /// </summary>
    bool PrintOutlinePages { get; set; }

    /// <summary>
    /// 获取或设置是否打印幻灯片页
    /// </summary>
    bool PrintSlidePages { get; set; }

    /// <summary>
    /// 应用打印选项
    /// </summary>
    void Apply();

    /// <summary>
    /// 重置打印选项
    /// </summary>
    void Reset();

    /// <summary>
    /// 获取打印选项信息
    /// </summary>
    /// <returns>打印选项信息字符串</returns>
    string GetPrintOptionsInfo();
}
