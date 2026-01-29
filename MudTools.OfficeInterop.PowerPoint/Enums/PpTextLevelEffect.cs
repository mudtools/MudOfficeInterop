//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定文本动画的级别效果。
/// </summary>
public enum PpTextLevelEffect
{
    /// <summary>
    /// 混合文本级别效果。
    /// </summary>
    ppAnimateLevelMixed = -2,

    /// <summary>
    /// 无文本级别动画。
    /// </summary>
    ppAnimateLevelNone = 0,

    /// <summary>
    /// 按第一级别（一级标题）设置文本动画。
    /// </summary>
    ppAnimateByFirstLevel = 1,

    /// <summary>
    /// 按第二级别（二级标题）设置文本动画。
    /// </summary>
    ppAnimateBySecondLevel = 2,

    /// <summary>
    /// 按第三级别（三级标题）设置文本动画。
    /// </summary>
    ppAnimateByThirdLevel = 3,

    /// <summary>
    /// 按第四级别（四级标题）设置文本动画。
    /// </summary>
    ppAnimateByFourthLevel = 4,

    /// <summary>
    /// 按第五级别（五级标题）设置文本动画。
    /// </summary>
    ppAnimateByFifthLevel = 5,

    /// <summary>
    /// 按所有级别设置文本动画。
    /// </summary>
    ppAnimateByAllLevels = 16
}