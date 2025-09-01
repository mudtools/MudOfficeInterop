//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定数据标签在图表中的位置
/// </summary>
public enum XlDataLabelPosition
{
    /// <summary>
    /// 数据标签位于中心位置
    /// </summary>
    xlLabelPositionCenter = -4108,
    
    /// <summary>
    /// 数据标签位于数据点上方
    /// </summary>
    xlLabelPositionAbove = 0,
    
    /// <summary>
    /// 数据标签位于数据点下方
    /// </summary>
    xlLabelPositionBelow = 1,
    
    /// <summary>
    /// 数据标签位于数据点左侧
    /// </summary>
    xlLabelPositionLeft = -4131,
    
    /// <summary>
    /// 数据标签位于数据点右侧
    /// </summary>
    xlLabelPositionRight = -4152,
    
    /// <summary>
    /// 数据标签位于数据系列末端外部
    /// </summary>
    xlLabelPositionOutsideEnd = 2,
    
    /// <summary>
    /// 数据标签位于数据系列末端内部
    /// </summary>
    xlLabelPositionInsideEnd = 3,
    
    /// <summary>
    /// 数据标签位于数据系列基部内部
    /// </summary>
    xlLabelPositionInsideBase = 4,
    
    /// <summary>
    /// 数据标签位置自动调整以获得最佳匹配
    /// </summary>
    xlLabelPositionBestFit = 5,
    
    /// <summary>
    /// 数据标签位置为混合模式
    /// </summary>
    xlLabelPositionMixed = 6,
    
    /// <summary>
    /// 数据标签位置为自定义
    /// </summary>
    xlLabelPositionCustom = 7
}