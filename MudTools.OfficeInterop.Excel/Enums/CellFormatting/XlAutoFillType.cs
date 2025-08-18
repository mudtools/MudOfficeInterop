//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 自动填充类型枚举
/// 用于指定使用自动填充功能时的填充方式
/// </summary>
public enum AutoFillType
{
    /// <summary>
    /// 复制填充
    /// 直接复制源单元格的内容和格式到目标区域
    /// </summary>
    xlFillCopy = 1,
    
    /// <summary>
    /// 按天填充
    /// 将日期序列按天递增填充
    /// </summary>
    xlFillDays = 5,
    
    /// <summary>
    /// 默认填充
    /// Excel根据选中区域的内容自动判断填充方式
    /// </summary>
    xlFillDefault = 0,
    
    /// <summary>
    /// 格式填充
    /// 只复制源单元格的格式，不复制内容
    /// </summary>
    xlFillFormats = 3,
    
    /// <summary>
    /// 按月填充
    /// 将日期序列按月递增填充
    /// </summary>
    xlFillMonths = 7,
    
    /// <summary>
    /// 序列填充
    /// 按照序列规律（如1,2,3或A,B,C）进行填充
    /// </summary>
    xlFillSeries = 2,
    
    /// <summary>
    /// 数值填充
    /// 只复制源单元格的数值，不复制公式
    /// </summary>
    xlFillValues = 4,
    
    /// <summary>
    /// 按工作日填充
    /// 将日期序列按工作日（周一至周五）填充
    /// </summary>
    xlFillWeekdays = 6,
    
    /// <summary>
    /// 按年填充
    /// 将日期序列按年递增填充
    /// </summary>
    xlFillYears = 8,
    
    /// <summary>
    /// 生长趋势填充
    /// 根据源数据的增长趋势进行填充
    /// </summary>
    xlGrowthTrend = 10,
    
    /// <summary>
    /// 线性趋势填充
    /// 根据源数据的线性趋势进行填充
    /// </summary>
    xlLinearTrend = 9,
    
    /// <summary>
    /// 快速填充
    /// 根据示例数据推断填充模式并应用到其他单元格
    /// </summary>
    xlFlashFill = 11
}