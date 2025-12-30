//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 指定 VBE 窗口类型
/// </summary>
public enum vbext_WindowType
{
    /// <summary>
    /// 代码窗口
    /// </summary>
    vbext_wt_CodeWindow = 0,

    /// <summary>
    /// 设计器窗口
    /// </summary>
    vbext_wt_Designer = 1,

    /// <summary>
    /// 对象浏览器窗口
    /// </summary>
    vbext_wt_Browser = 2,

    /// <summary>
    /// 监视窗口
    /// </summary>
    vbext_wt_Watch = 3,

    /// <summary>
    /// 本地窗口
    /// </summary>
    vbext_wt_Locals = 4,

    /// <summary>
    /// 立即窗口
    /// </summary>
    vbext_wt_Immediate = 5,

    /// <summary>
    /// 项目窗口
    /// </summary>
    vbext_wt_ProjectWindow = 6,

    /// <summary>
    /// 属性窗口
    /// </summary>
    vbext_wt_PropertyWindow = 7,

    /// <summary>
    /// 查找窗口
    /// </summary>
    vbext_wt_Find = 8,

    /// <summary>
    /// 查找/替换窗口
    /// </summary>
    vbext_wt_FindReplace = 9,

    /// <summary>
    /// 工具箱窗口
    /// </summary>
    vbext_wt_Toolbox = 10,

    /// <summary>
    /// 链接窗口框架
    /// </summary>
    vbext_wt_LinkedWindowFrame = 11,

    /// <summary>
    /// 主窗口
    /// </summary>
    vbext_wt_MainWindow = 12,

    /// <summary>
    /// 工具窗口
    /// </summary>
    vbext_wt_ToolWindow = 15
}