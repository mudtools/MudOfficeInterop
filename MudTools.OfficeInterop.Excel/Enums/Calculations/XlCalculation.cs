//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 计算模式枚举
/// 用于控制Excel工作表的计算方式
/// </summary>
public enum XlCalculation
{
    /// <summary>
    /// 自动计算模式
    /// Excel会在工作表数据发生更改时自动重新计算相关公式
    /// </summary>
    xlCalculationAutomatic = -4105,
    
    /// <summary>
    /// 手动计算模式
    /// Excel只在用户明确要求时才重新计算公式（按F9键）
    /// </summary>
    xlCalculationManual = -4135,
    
    /// <summary>
    /// 半自动计算模式
    /// Excel会自动重新计算工作表中的公式，但不会重新计算其他工作表中的相关公式
    /// </summary>
    xlCalculationSemiautomatic = 2
}
