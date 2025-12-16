//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定图表中图片标记的放置位置
/// </summary>
public enum XlChartPicturePlacement
{
    /// <summary>
    /// 图片标记显示在数据标记的两侧
    /// </summary>
    xlSides = 1,

    /// <summary>
    /// 图片标记显示在数据标记的末端
    /// </summary>
    xlEnd,

    /// <summary>
    /// 图片标记显示在数据标记的末端和两侧
    /// </summary>
    xlEndSides,

    /// <summary>
    /// 图片标记显示在数据标记的前端
    /// </summary>
    xlFront,

    /// <summary>
    /// 图片标记显示在数据标记的前端和两侧
    /// </summary>
    xlFrontSides,

    /// <summary>
    /// 图片标记显示在数据标记的前端和末端
    /// </summary>
    xlFrontEnd,

    /// <summary>
    /// 图片标记显示在数据标记的所有面上
    /// </summary>
    xlAllFaces
}