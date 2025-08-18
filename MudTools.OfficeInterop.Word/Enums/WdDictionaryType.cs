//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 字典类型枚举
/// 用于指定不同类型的词典，如拼写检查、语法检查等
/// </summary>
public enum WdDictionaryType
{
    /// <summary>
    /// 拼写检查字典
    /// </summary>
    wdSpelling,
    
    /// <summary>
    /// 语法检查字典
    /// </summary>
    wdGrammar,
    
    /// <summary>
    /// 同义词库字典
    /// </summary>
    wdThesaurus,
    
    /// <summary>
    /// 断字处理字典
    /// </summary>
    wdHyphenation,
    
    /// <summary>
    /// 完整拼写检查字典
    /// </summary>
    wdSpellingComplete,
    
    /// <summary>
    /// 自定义拼写检查字典
    /// </summary>
    wdSpellingCustom,
    
    /// <summary>
    /// 法律术语拼写检查字典
    /// </summary>
    wdSpellingLegal,
    
    /// <summary>
    /// 医学术语拼写检查字典
    /// </summary>
    wdSpellingMedical,
    
    /// <summary>
    /// 韩文汉字转换字典
    /// </summary>
    wdHangulHanjaConversion,
    
    /// <summary>
    /// 自定义韩文汉字转换字典
    /// </summary>
    wdHangulHanjaConversionCustom
}