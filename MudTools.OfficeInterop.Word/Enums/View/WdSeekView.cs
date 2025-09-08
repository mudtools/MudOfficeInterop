//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在页面视图中要显示的文档元素类型
/// 用于控制在Word页面视图中查看文档的不同部分，如页眉、页脚、脚注等
/// </summary>
public enum WdSeekView
{
    /// <summary>
    /// 主文档视图
    /// </summary>
    wdSeekMainDocument,
    
    /// <summary>
    /// 主页眉视图
    /// </summary>
    wdSeekPrimaryHeader,
    
    /// <summary>
    /// 首页页眉视图
    /// </summary>
    wdSeekFirstPageHeader,
    
    /// <summary>
    /// 偶数页页眉视图
    /// </summary>
    wdSeekEvenPagesHeader,
    
    /// <summary>
    /// 主页脚视图
    /// </summary>
    wdSeekPrimaryFooter,
    
    /// <summary>
    /// 首页页脚视图
    /// </summary>
    wdSeekFirstPageFooter,
    
    /// <summary>
    /// 偶数页页脚视图
    /// </summary>
    wdSeekEvenPagesFooter,
    
    /// <summary>
    /// 脚注视图
    /// </summary>
    wdSeekFootnotes,
    
    /// <summary>
    /// 尾注视图
    /// </summary>
    wdSeekEndnotes,
    
    /// <summary>
    /// 当前页页眉视图
    /// </summary>
    wdSeekCurrentPageHeader,
    
    /// <summary>
    /// 当前页页脚视图
    /// </summary>
    wdSeekCurrentPageFooter
}