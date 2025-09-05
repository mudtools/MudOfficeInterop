//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档部分类型枚举
/// 用于指定文档中的不同部分（如正文、页眉、页脚、脚注等）
/// </summary>
public enum WdStoryType
{
    /// <summary>
    /// 主文本部分
    /// </summary>
    wdMainTextStory = 1,
    
    /// <summary>
    /// 脚注部分
    /// </summary>
    wdFootnotesStory,
    
    /// <summary>
    /// 尾注部分
    /// </summary>
    wdEndnotesStory,
    
    /// <summary>
    /// 批注部分
    /// </summary>
    wdCommentsStory,
    
    /// <summary>
    /// 文本框部分
    /// </summary>
    wdTextFrameStory,
    
    /// <summary>
    /// 偶数页页眉部分
    /// </summary>
    wdEvenPagesHeaderStory,
    
    /// <summary>
    /// 主页眉部分
    /// </summary>
    wdPrimaryHeaderStory,
    
    /// <summary>
    /// 偶数页页脚部分
    /// </summary>
    wdEvenPagesFooterStory,
    
    /// <summary>
    /// 主页脚部分
    /// </summary>
    wdPrimaryFooterStory,
    
    /// <summary>
    /// 首页页眉部分
    /// </summary>
    wdFirstPageHeaderStory,
    
    /// <summary>
    /// 首页页脚部分
    /// </summary>
    wdFirstPageFooterStory,
    
    /// <summary>
    /// 脚注分隔符部分
    /// </summary>
    wdFootnoteSeparatorStory,
    
    /// <summary>
    /// 脚注延续分隔符部分
    /// </summary>
    wdFootnoteContinuationSeparatorStory,
    
    /// <summary>
    /// 脚注延续通知部分
    /// </summary>
    wdFootnoteContinuationNoticeStory,
    
    /// <summary>
    /// 尾注分隔符部分
    /// </summary>
    wdEndnoteSeparatorStory,
    
    /// <summary>
    /// 尾注延续分隔符部分
    /// </summary>
    wdEndnoteContinuationSeparatorStory,
    
    /// <summary>
    /// 尾注延续通知部分
    /// </summary>
    wdEndnoteContinuationNoticeStory
}