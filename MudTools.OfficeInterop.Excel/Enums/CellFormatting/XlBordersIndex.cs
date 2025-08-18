//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 边框索引枚举
/// 用于指定单元格或单元格区域中特定的边框位置
/// </summary>
public enum XlBordersIndex
{
    /// <summary>
    /// 内部水平边框
    /// 选定区域内部的水平边框线
    /// </summary>
    xlInsideHorizontal = 12,
    
    /// <summary>
    /// 内部垂直边框
    /// 选定区域内部的垂直边框线
    /// </summary>
    xlInsideVertical = 11,
    
    /// <summary>
    /// 向下对角线边框
    /// 从左上到右下的对角线边框
    /// </summary>
    xlDiagonalDown = 5,
    
    /// <summary>
    /// 向上对角线边框
    /// 从左下到右上的对角线边框
    /// </summary>
    xlDiagonalUp = 6,
    
    /// <summary>
    /// 底部边框
    /// 单元格或区域的底部边框线
    /// </summary>
    xlEdgeBottom = 9,
    
    /// <summary>
    /// 左侧边框
    /// 单元格或区域的左侧边框线
    /// </summary>
    xlEdgeLeft = 7,
    
    /// <summary>
    /// 右侧边框
    /// 单元格或区域的右侧边框线
    /// </summary>
    xlEdgeRight = 10,
    
    /// <summary>
    /// 顶部边框
    /// 单元格或区域的顶部边框线
    /// </summary>
    xlEdgeTop = 8
}