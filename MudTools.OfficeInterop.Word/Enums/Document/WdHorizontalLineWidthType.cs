//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Word 如何解释指定水平线的宽度（长度）
/// </summary>
public enum WdHorizontalLineWidthType
{
    /// <summary>
    /// Word 将指定水平线的宽度（长度）解释为屏幕宽度的百分比。这是使用 AddHorizontalLineStandard 方法添加水平线时的默认值。设置水平线的 PercentWidth 属性会将 WidthType 属性设置为此值。
    /// </summary>
    wdHorizontalLinePercentWidth = -1,

    /// <summary>
    /// Microsoft Word 将指定水平线的宽度（长度）解释为固定值（以磅为单位）。这是使用 AddHorizontalLine 方法添加水平线时的默认值。设置与水平线关联的 InlineShape 对象的 Width 属性会将 WidthType 属性设置为此值。
    /// </summary>
    wdHorizontalLineFixedWidth = -2
}