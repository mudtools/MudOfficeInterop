//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定打印纸张来源的纸盒类型
/// </summary>
public enum WdPaperTray
{
    /// <summary>
    /// 打印机默认纸盒
    /// </summary>
    wdPrinterDefaultBin = 0,
    
    /// <summary>
    /// 打印机上层纸盒
    /// </summary>
    wdPrinterUpperBin = 1,
    
    /// <summary>
    /// 打印机唯一纸盒
    /// </summary>
    wdPrinterOnlyBin = 1,
    
    /// <summary>
    /// 打印机下层纸盒
    /// </summary>
    wdPrinterLowerBin = 2,
    
    /// <summary>
    /// 打印机中层纸盒
    /// </summary>
    wdPrinterMiddleBin = 3,
    
    /// <summary>
    /// 手动送纸
    /// </summary>
    wdPrinterManualFeed = 4,
    
    /// <summary>
    /// 信封送纸
    /// </summary>
    wdPrinterEnvelopeFeed = 5,
    
    /// <summary>
    /// 手动信封送纸
    /// </summary>
    wdPrinterManualEnvelopeFeed = 6,
    
    /// <summary>
    /// 自动纸张送纸
    /// </summary>
    wdPrinterAutomaticSheetFeed = 7,
    
    /// <summary>
    /// 拖纸器送纸
    /// </summary>
    wdPrinterTractorFeed = 8,
    
    /// <summary>
    /// 小格式纸盒
    /// </summary>
    wdPrinterSmallFormatBin = 9,
    
    /// <summary>
    /// 大格式纸盒
    /// </summary>
    wdPrinterLargeFormatBin = 10,
    
    /// <summary>
    /// 大容量纸盒
    /// </summary>
    wdPrinterLargeCapacityBin = 11,
    
    /// <summary>
    /// 纸盒供纸
    /// </summary>
    wdPrinterPaperCassette = 14,
    
    /// <summary>
    /// 表单来源
    /// </summary>
    wdPrinterFormSource = 15
}