//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 引用样式枚举
/// 用于指定单元格引用的样式格式
/// </summary>
public enum XlReferenceStyle
{
    /// <summary>
    /// A1引用样式
    /// 使用字母表示列（A, B, C...），数字表示行（1, 2, 3...）的引用样式，如A1, B2
    /// </summary>
    xlA1 = 1,
    
    /// <summary>
    /// R1C1引用样式
    /// 使用R表示行，C表示列的引用样式，如R1C1, R2C3
    /// </summary>
    xlR1C1 = -4150
}