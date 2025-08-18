//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文本方向枚举
/// 用于指定文本的排列方向
/// </summary>
public enum WdTextOrientation
{
    /// <summary>
    /// 水平文本方向（从左到右）
    /// </summary>
    wdTextOrientationHorizontal = 0,
    
    /// <summary>
    /// 向上文本方向（从下到上）
    /// </summary>
    wdTextOrientationUpward = 2,
    
    /// <summary>
    /// 向下文本方向（从上到下）
    /// </summary>
    wdTextOrientationDownward = 3,
    
    /// <summary>
    /// 纵向远东文本方向（垂直排列，适用于中文、日文等）
    /// </summary>
    wdTextOrientationVerticalFarEast = 1,
    
    /// <summary>
    /// 水平旋转远东文本方向（水平排列但字符旋转90度，适用于中文、日文等）
    /// </summary>
    wdTextOrientationHorizontalRotatedFarEast = 4
}