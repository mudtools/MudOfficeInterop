//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 窗口类型枚举
/// 用于指定Excel中不同类型的窗口
/// </summary>
public enum XlWindowType
{
    /// <summary>
    /// 图表作为独立窗口
    /// 图表以独立窗口的形式显示
    /// </summary>
    xlChartAsWindow = 5,
    
    /// <summary>
    /// 图表嵌入原位置
    /// 图表嵌入在工作表中的原位置显示
    /// </summary>
    xlChartInPlace = 4,
    
    /// <summary>
    /// 剪贴板窗口
    /// 显示剪贴板内容的窗口
    /// </summary>
    xlClipboard = 3,
    
    /// <summary>
    /// 信息窗口
    /// 显示相关信息的窗口
    /// </summary>
    xlInfo = -4129,
    
    /// <summary>
    /// 工作簿窗口
    /// 显示工作簿内容的主窗口
    /// </summary>
    xlWorkbook = 1
}