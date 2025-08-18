//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 打印选项实现类
/// </summary>
internal class PowerPointPrintOptions : IPowerPointPrintOptions
{
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置打印份数
    /// </summary>
    public int Copies { get; set; } = 1;

    /// <summary>
    /// 获取或设置是否逐份打印
    /// </summary>
    public bool Collate { get; set; } = true;

    /// <summary>
    /// 获取或设置打印范围
    /// </summary>
    public string PrintRange { get; set; } = "";

    /// <summary>
    /// 获取或设置打印到文件
    /// </summary>
    public bool PrintToFile { get; set; } = false;

    /// <summary>
    /// 获取或设置输出文件名
    /// </summary>
    public string OutputFileName { get; set; } = "";

    /// <summary>
    /// 获取或设置打印机名称
    /// </summary>
    public string PrinterName { get; set; } = "";

    /// <summary>
    /// 获取或设置打印质量
    /// </summary>
    public int PrintQuality { get; set; } = 600;

    /// <summary>
    /// 获取或设置打印颜色模式
    /// </summary>
    public int ColorMode { get; set; } = 2; // 彩色

    /// <summary>
    /// 获取或设置打印方向
    /// </summary>
    public int Orientation { get; set; } = 1; // 纵向

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    public int PaperSize { get; set; } = 9; // A4

    /// <summary>
    /// 获取或设置纸张来源
    /// </summary>
    public int PaperSource { get; set; } = 7; // 自动

    /// <summary>
    /// 获取或设置打印分辨率
    /// </summary>
    public int Resolution { get; set; } = 600;

    /// <summary>
    /// 获取或设置打印份数
    /// </summary>
    public int NumberOfCopies { get; set; } = 1;

    /// <summary>
    /// 获取或设置是否双面打印
    /// </summary>
    public bool Duplex { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印注释
    /// </summary>
    public bool PrintComments { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印隐藏幻灯片
    /// </summary>
    public bool PrintHiddenSlides { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印幻灯片编号
    /// </summary>
    public bool PrintSlideNumbers { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印日期和时间
    /// </summary>
    public bool PrintDateTime { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印页眉页脚
    /// </summary>
    public bool PrintHeadersFooters { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印备注
    /// </summary>
    public bool PrintNotes { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印讲义
    /// </summary>
    public bool PrintHandouts { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印大纲
    /// </summary>
    public bool PrintOutline { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印幻灯片
    /// </summary>
    public bool PrintSlides { get; set; } = true;

    /// <summary>
    /// 获取或设置是否打印备注页
    /// </summary>
    public bool PrintNotesPages { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印讲义页
    /// </summary>
    public bool PrintHandoutPages { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印大纲页
    /// </summary>
    public bool PrintOutlinePages { get; set; } = false;

    /// <summary>
    /// 获取或设置是否打印幻灯片页
    /// </summary>
    public bool PrintSlidePages { get; set; } = true;

    /// <summary>
    /// 构造函数
    /// </summary>
    internal PowerPointPrintOptions()
    {
        _disposedValue = false;
    }

    /// <summary>
    /// 应用打印选项
    /// </summary>
    public void Apply()
    {
        try
        {
            // 打印选项应用需要通过演示文稿对象实现
            throw new NotImplementedException("Applying print options is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply print options.", ex);
        }
    }

    /// <summary>
    /// 重置打印选项
    /// </summary>
    public void Reset()
    {
        try
        {
            Copies = 1;
            Collate = true;
            PrintRange = "";
            PrintToFile = false;
            OutputFileName = "";
            PrinterName = "";
            PrintQuality = 600;
            ColorMode = 2;
            Orientation = 1;
            PaperSize = 9;
            PaperSource = 7;
            Resolution = 600;
            NumberOfCopies = 1;
            Duplex = false;
            PrintComments = false;
            PrintHiddenSlides = false;
            PrintSlideNumbers = false;
            PrintDateTime = false;
            PrintHeadersFooters = false;
            PrintNotes = false;
            PrintHandouts = false;
            PrintOutline = false;
            PrintSlides = true;
            PrintNotesPages = false;
            PrintHandoutPages = false;
            PrintOutlinePages = false;
            PrintSlidePages = true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset print options.", ex);
        }
    }

    /// <summary>
    /// 获取打印选项信息
    /// </summary>
    /// <returns>打印选项信息字符串</returns>
    public string GetPrintOptionsInfo()
    {
        try
        {
            return $"PrintOptions - Copies: {Copies}, Collate: {Collate}, Printer: {PrinterName}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get print options info.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}