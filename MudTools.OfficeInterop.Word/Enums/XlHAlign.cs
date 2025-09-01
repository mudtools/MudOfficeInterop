//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Excel中的水平对齐方式
/// </summary>
public enum XlHAlign
{
    /// <summary>
    /// 居中对齐
    /// </summary>
    xlHAlignCenter = -4108,
    
    /// <summary>
    /// 跨选定区域居中对齐
    /// </summary>
    xlHAlignCenterAcrossSelection = 7,
    
    /// <summary>
    /// 分散对齐
    /// </summary>
    xlHAlignDistributed = -4117,
    
    /// <summary>
    /// 填充对齐
    /// </summary>
    xlHAlignFill = 5,
    
    /// <summary>
    /// 常规对齐
    /// </summary>
    xlHAlignGeneral = 1,
    
    /// <summary>
    /// 两端对齐
    /// </summary>
    xlHAlignJustify = -4130,
    
    /// <summary>
    /// 左对齐
    /// </summary>
    xlHAlignLeft = -4131,
    
    /// <summary>
    /// 右对齐
    /// </summary>
    xlHAlignRight = -4152
}