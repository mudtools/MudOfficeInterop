//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定自定义邮件标签的页面尺寸
/// </summary>
[Guid("B116B479-488C-3D69-BFBE-A64DD14F3BB9")]
public enum WdCustomLabelPageSize
{
    /// <summary>
    /// 标准信函纵向标签尺寸
    /// </summary>
    wdCustomLabelLetter,

    /// <summary>
    /// 标准信函横向标签尺寸
    /// </summary>
    wdCustomLabelLetterLS,

    /// <summary>
    /// A4纵向标签尺寸
    /// </summary>
    wdCustomLabelA4,

    /// <summary>
    /// A4横向标签尺寸
    /// </summary>
    wdCustomLabelA4LS,

    /// <summary>
    /// A5纵向标签尺寸
    /// </summary>
    wdCustomLabelA5,

    /// <summary>
    /// A5横向标签尺寸
    /// </summary>
    wdCustomLabelA5LS,

    /// <summary>
    /// B5标签尺寸
    /// </summary>
    wdCustomLabelB5,

    /// <summary>
    /// 迷你标签尺寸
    /// </summary>
    wdCustomLabelMini,

    /// <summary>
    /// 连续折页标签尺寸
    /// </summary>
    wdCustomLabelFanfold,

    /// <summary>
    /// 半页纵向标签尺寸
    /// </summary>
    wdCustomLabelVertHalfSheet,

    /// <summary>
    /// 半页横向标签尺寸
    /// </summary>
    wdCustomLabelVertHalfSheetLS,

    /// <summary>
    /// 日式信封纵向标签尺寸
    /// </summary>
    wdCustomLabelHigaki,

    /// <summary>
    /// 日式信封横向标签尺寸
    /// </summary>
    wdCustomLabelHigakiLS,

    /// <summary>
    /// B4 JIS标签尺寸
    /// </summary>
    wdCustomLabelB4JIS
}