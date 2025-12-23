//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定添加到指定范围的拼音文本的对齐方式
/// </summary>
public enum WdPhoneticGuideAlignmentType
{
    /// <summary>
    /// Microsoft Word 将拼音文本在指定范围上方居中显示。这是默认值。
    /// </summary>
    wdPhoneticGuideAlignmentCenter,

    /// <summary>
    /// Word 按照 0:1:0 的比例调整拼音文本的内部和外部间距。
    /// </summary>
    wdPhoneticGuideAlignmentZeroOneZero,

    /// <summary>
    /// Word 按照 1:2:1 的比例调整拼音文本的内部和外部间距。
    /// </summary>
    wdPhoneticGuideAlignmentOneTwoOne,

    /// <summary>
    /// Word 将拼音文本左对齐到指定范围。
    /// </summary>
    wdPhoneticGuideAlignmentLeft,

    /// <summary>
    /// Word 将拼音文本右对齐到指定范围。
    /// </summary>
    wdPhoneticGuideAlignmentRight,

    /// <summary>
    /// Word 将拼音文本在垂直文本的右侧对齐。
    /// </summary>
    wdPhoneticGuideAlignmentRightVertical
}