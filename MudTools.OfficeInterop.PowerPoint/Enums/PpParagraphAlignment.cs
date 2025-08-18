//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 段落对齐方式枚举
/// </summary>
public enum PpParagraphAlignment
{
    /// <summary>
    /// 混合对齐方式
    /// </summary>
    ppAlignmentMixed = -2,
    
    /// <summary>
    /// 左对齐
    /// </summary>
    ppAlignLeft = 1,
    
    /// <summary>
    /// 居中对齐
    /// </summary>
    ppAlignCenter = 2,
    
    /// <summary>
    /// 右对齐
    /// </summary>
    ppAlignRight = 3,
    
    /// <summary>
    /// 两端对齐
    /// </summary>
    ppAlignJustify = 4,
    
    /// <summary>
    /// 分散对齐
    /// </summary>
    ppAlignDistribute = 5,
    
    /// <summary>
    /// 泰语分散对齐
    /// </summary>
    ppAlignThaiDistribute = 6,
    
    /// <summary>
    /// 低度两端对齐
    /// </summary>
    ppAlignJustifyLow = 7
}