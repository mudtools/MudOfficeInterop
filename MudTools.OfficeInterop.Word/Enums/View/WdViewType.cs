//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档视图类型枚举
/// 用于指定文档的显示视图
/// </summary>
public enum WdViewType
{
    /// <summary>
    /// 普通视图 - 显示文档的基本格式
    /// </summary>
    wdNormalView = 1,
    
    /// <summary>
    /// 大纲视图 - 显示文档的结构和层次
    /// </summary>
    wdOutlineView = 2,
    
    /// <summary>
    /// 页面视图 - 显示文档的打印效果
    /// </summary>
    wdPrintView = 3,
    
    /// <summary>
    /// Web 版式视图 - 以网页形式显示文档
    /// </summary>
    wdWebView = 4,
    
    /// <summary>
    /// 阅读视图 - 优化阅读体验的视图模式
    /// </summary>
    wdReadingView = 5,
    
    /// <summary>
    /// 主控文档视图 - 用于处理主控文档和子文档
    /// </summary>
    wdMasterView = 6
}