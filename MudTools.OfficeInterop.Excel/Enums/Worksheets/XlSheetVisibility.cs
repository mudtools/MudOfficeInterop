//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 工作表可见性枚举
/// 用于控制工作表的可见性状态
/// </summary>
public enum XlSheetVisibility
{
    /// <summary>
    /// 工作表可见
    /// 用户可以通过工作表标签访问该工作表
    /// </summary>
    xlSheetVisible = -1,
    
    /// <summary>
    /// 工作表隐藏
    /// 工作表被隐藏，但可以通过"取消隐藏"命令恢复显示
    /// </summary>
    xlSheetHidden = 0,
    
    /// <summary>
    /// 工作表深度隐藏
    /// 工作表被隐藏且无法通过Excel界面中的"取消隐藏"命令恢复显示，只能通过编程方式或VBA代码取消隐藏
    /// </summary>
    xlSheetVeryHidden = 2
}