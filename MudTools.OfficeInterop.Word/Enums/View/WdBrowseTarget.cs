//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在Word文档中浏览时要定位的目标项类型
/// </summary>
public enum WdBrowseTarget
{
    /// <summary>
    /// 浏览页面
    /// </summary>
    wdBrowsePage = 1,
    
    /// <summary>
    /// 浏览章节
    /// </summary>
    wdBrowseSection,
    
    /// <summary>
    /// 浏览批注
    /// </summary>
    wdBrowseComment,
    
    /// <summary>
    /// 浏览脚注
    /// </summary>
    wdBrowseFootnote,
    
    /// <summary>
    /// 浏览尾注
    /// </summary>
    wdBrowseEndnote,
    
    /// <summary>
    /// 浏览域
    /// </summary>
    wdBrowseField,
    
    /// <summary>
    /// 浏览表格
    /// </summary>
    wdBrowseTable,
    
    /// <summary>
    /// 浏览图形
    /// </summary>
    wdBrowseGraphic,
    
    /// <summary>
    /// 浏览标题
    /// </summary>
    wdBrowseHeading,
    
    /// <summary>
    /// 浏览编辑处
    /// </summary>
    wdBrowseEdit,
    
    /// <summary>
    /// 浏览查找内容
    /// </summary>
    wdBrowseFind,
    
    /// <summary>
    /// 浏览转到目标
    /// </summary>
    wdBrowseGoTo
}