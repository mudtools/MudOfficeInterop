//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 图例位置枚举
/// 用于指定图表中图例的位置
/// </summary>
public enum XlLegendPosition
{
    /// <summary>
    /// 底部
    /// 图例显示在图表底部
    /// </summary>
    xlLegendPositionBottom = -4107,
    
    /// <summary>
    /// 角落
    /// 图例显示在图表角落（通常在图表内部）
    /// </summary>
    xlLegendPositionCorner = 2,
    
    /// <summary>
    /// 左侧
    /// 图例显示在图表左侧
    /// </summary>
    xlLegendPositionLeft = -4131,
    
    /// <summary>
    /// 右侧
    /// 图例显示在图表右侧
    /// </summary>
    xlLegendPositionRight = -4152,
    
    /// <summary>
    /// 顶部
    /// 图例显示在图表顶部
    /// </summary>
    xlLegendPositionTop = -4160,
    
    /// <summary>
    /// 自定义位置
    /// 图例位置由用户自定义
    /// </summary>
    xlLegendPositionCustom = -4161
}
