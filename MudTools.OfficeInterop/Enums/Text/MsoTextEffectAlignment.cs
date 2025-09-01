//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定艺术字文本的对齐方式
/// </summary>
public enum MsoTextEffectAlignment
{
    /// <summary>
    /// 混合对齐方式（通常用于表示未设置或多种方式的组合）
    /// </summary>
    msoTextEffectAlignmentMixed = -2,
    
    /// <summary>
    /// 左对齐
    /// </summary>
    msoTextEffectAlignmentLeft = 1,
    
    /// <summary>
    /// 居中对齐
    /// </summary>
    msoTextEffectAlignmentCentered = 2,
    
    /// <summary>
    /// 右对齐
    /// </summary>
    msoTextEffectAlignmentRight = 3,
    
    /// <summary>
    /// 字母间距调整对齐（分散对齐，调整字母间距以填充行宽）
    /// </summary>
    msoTextEffectAlignmentLetterJustify = 4,
    
    /// <summary>
    /// 单词间距调整对齐（调整单词间距以填充行宽）
    /// </summary>
    msoTextEffectAlignmentWordJustify = 5,
    
    /// <summary>
    /// 拉伸调整对齐（通过拉伸或压缩文本以填充行宽）
    /// </summary>
    msoTextEffectAlignmentStretchJustify = 6
}