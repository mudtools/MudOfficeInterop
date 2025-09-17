//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在Microsoft Word中使用的度量单位类型
/// </summary>
public enum WdUnits
{
    /// <summary>
    /// 字符
    /// </summary>
    wdCharacter = 1,
    
    /// <summary>
    /// 字
    /// </summary>
    wdWord,
    
    /// <summary>
    /// 句子
    /// </summary>
    wdSentence,
    
    /// <summary>
    /// 段落
    /// </summary>
    wdParagraph,
    
    /// <summary>
    /// 行
    /// </summary>
    wdLine,
    
    /// <summary>
    /// 部分
    /// </summary>
    wdStory,
    
    /// <summary>
    /// 屏幕尺寸
    /// </summary>
    wdScreen,
    
    /// <summary>
    /// 节
    /// </summary>
    wdSection,
    
    /// <summary>
    /// 列
    /// </summary>
    wdColumn,
    
    /// <summary>
    /// 行
    /// </summary>
    wdRow,
    
    /// <summary>
    /// 窗口
    /// </summary>
    wdWindow,
    
    /// <summary>
    /// 单元格
    /// </summary>
    wdCell,
    
    /// <summary>
    /// 字符格式
    /// </summary>
    wdCharacterFormatting,
    
    /// <summary>
    /// 段落格式
    /// </summary>
    wdParagraphFormatting,
    
    /// <summary>
    /// 表格
    /// </summary>
    wdTable,
    
    /// <summary>
    /// 所选项
    /// </summary>
    wdItem
}