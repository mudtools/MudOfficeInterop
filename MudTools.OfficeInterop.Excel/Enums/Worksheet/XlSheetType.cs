//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 工作表类型枚举
/// 用于指定工作簿中工作表的类型
/// </summary>
public enum XlSheetType
{
    /// <summary>
    /// 图表工作表
    /// 仅包含图表的独立工作表
    /// </summary>
    xlChart = -4109,
    
    /// <summary>
    /// 对话框工作表
    /// 包含对话框元素的工作表（较早版本Excel的功能）
    /// </summary>
    xlDialogSheet = -4116,
    
    /// <summary>
    /// Excel 4 国际宏工作表
    /// Excel 4.0版本的国际宏工作表
    /// </summary>
    xlExcel4IntlMacroSheet = 4,
    
    /// <summary>
    /// Excel 4 宏工作表
    /// Excel 4.0版本的宏工作表
    /// </summary>
    xlExcel4MacroSheet = 3,
    
    /// <summary>
    /// 普通工作表
    /// 包含单元格网格的标准工作表
    /// </summary>
    xlWorksheet = -4167
}