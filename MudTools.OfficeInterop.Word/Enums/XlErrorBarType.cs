//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定图表误差线的类型
/// 用于定义Excel图表中数据系列误差线的计算方式
/// </summary>
public enum XlErrorBarType
{
    /// <summary>
    /// 自定义误差线类型
    /// 用户可以指定正负误差量的具体值
    /// </summary>
    xlErrorBarTypeCustom = -4114,
    
    /// <summary>
    /// 固定值误差线类型
    /// 误差量为固定的数值
    /// </summary>
    xlErrorBarTypeFixedValue = 1,
    
    /// <summary>
    /// 百分比误差线类型
    /// 误差量为数据点值的百分比
    /// </summary>
    xlErrorBarTypePercent = 2,
    
    /// <summary>
    /// 标准偏差误差线类型
    /// 误差量为数据的标准偏差
    /// </summary>
    xlErrorBarTypeStDev = -4155,
    
    /// <summary>
    /// 标准误差线类型
    /// 误差量为数据的标准误差
    /// </summary>
    xlErrorBarTypeStError = 4
}