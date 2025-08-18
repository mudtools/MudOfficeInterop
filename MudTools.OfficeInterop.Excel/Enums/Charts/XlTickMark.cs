//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 刻度标记枚举
/// 用于指定图表坐标轴上刻度线的类型和位置
/// </summary>
public enum XlTickMark
{
    /// <summary>
    /// 交叉刻度线
    /// 刻度线穿过坐标轴
    /// </summary>
    xlTickMarkCross = 4,
    
    /// <summary>
    /// 内部刻度线
    /// 刻度线向坐标轴内部延伸
    /// </summary>
    xlTickMarkInside = 2,
    
    /// <summary>
    /// 无刻度线
    /// 不显示刻度线
    /// </summary>
    xlTickMarkNone = -4142,
    
    /// <summary>
    /// 外部刻度线
    /// 刻度线向坐标轴外部延伸
    /// </summary>
    xlTickMarkOutside = 3
}