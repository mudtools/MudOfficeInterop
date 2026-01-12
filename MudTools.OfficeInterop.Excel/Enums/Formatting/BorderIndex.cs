//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 边框索引枚举
/// 用于指定单元格边框的不同位置
/// </summary>
public enum BorderIndex
{
    /// <summary>
    /// 左边框
    /// </summary>
    Left = 7,
    
    /// <summary>
    /// 右边框
    /// </summary>
    Right = 10,
    
    /// <summary>
    /// 上边框
    /// </summary>
    Top = 8,
    
    /// <summary>
    /// 下边框
    /// </summary>
    Bottom = 9,
    
    /// <summary>
    /// 向下对角线边框
    /// </summary>
    DiagonalDown = 5,
    
    /// <summary>
    /// 向上对角线边框
    /// </summary>
    DiagonalUp = 6
}