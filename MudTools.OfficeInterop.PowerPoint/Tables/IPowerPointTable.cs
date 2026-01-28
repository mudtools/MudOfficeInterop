//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

using System;

/// <summary>
/// 表示 PowerPoint 中的一个表格。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTable : IDisposable
{
    /// <summary>
    /// 获取创建此表格的 PowerPoint 应用程序对象。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此表格的父对象。
    /// </summary>
    /// <value>表示父对象的 <see cref="object"/>。</value>
    object Parent { get; }

    /// <summary>
    /// 获取表格的列集合。
    /// </summary>
    /// <value>表示表格列集合的 <see cref="IPowerPointColumns"/> 对象。</value>
    [ComPropertyWrap(NeedConvert = true)]
    IPowerPointColumns? Columns { get; }

    /// <summary>
    /// 获取表格的行集合。
    /// </summary>
    /// <value>表示表格行集合的 <see cref="IPowerPointRows"/> 对象。</value>
    [ComPropertyWrap(NeedConvert = true)]
    IPowerPointRows? Rows { get; }

    /// <summary>
    /// 根据指定的行和列索引获取表格中的单元格。
    /// </summary>
    /// <param name="row">要获取的单元格所在的行索引（从 1 开始）。</param>
    /// <param name="column">要获取的单元格所在的列索引（从 1 开始）。</param>
    /// <returns>指定位置的 <see cref="IPowerPointCell"/> 对象。</returns>
    IPowerPointCell? Cell(int row, int column);

    /// <summary>
    /// 获取或设置表格的方向。
    /// </summary>
    /// <value>表示表格方向的 <see cref="PpDirection"/> 枚举值。</value>
    PpDirection TableDirection { get; set; }

    /// <summary>
    /// 合并表格中相邻单元格的边框，使表格外观更简洁。
    /// </summary>
    void MergeBorders();

    /// <summary>
    /// 获取或设置一个值，该值指示是否为第一行应用特殊格式。
    /// </summary>
    /// <value>如果为第一行应用特殊格式，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool FirstRow { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否为最后一行应用特殊格式。
    /// </summary>
    /// <value>如果为最后一行应用特殊格式，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool LastRow { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否为第一列应用特殊格式。
    /// </summary>
    /// <value>如果为第一列应用特殊格式，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool FirstCol { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否为最后一列应用特殊格式。
    /// </summary>
    /// <value>如果为最后一列应用特殊格式，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool LastCol { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否启用水平条纹（行交替样式）。
    /// </summary>
    /// <value>如果启用水平条纹，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool HorizBanding { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否启用垂直条纹（列交替样式）。
    /// </summary>
    /// <value>如果启用垂直条纹，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool VertBanding { get; set; }

    /// <summary>
    /// 获取表格的样式对象。
    /// </summary>
    /// <value>表示表格样式的 <see cref="IPowerPointTableStyle"/> 对象。</value>
    IPowerPointTableStyle? Style { get; }

    /// <summary>
    /// 获取表格的背景对象。
    /// </summary>
    /// <value>表示表格背景的 <see cref="IPowerPointTableBackground"/> 对象。</value>
    IPowerPointTableBackground? Background { get; }

    /// <summary>
    /// 按指定比例等比例缩放表格。
    /// </summary>
    /// <param name="scale">缩放比例，例如 1.5 表示放大 150%，0.5 表示缩小 50%。</param>
    void ScaleProportionally(float scale);

    /// <summary>
    /// 对表格应用指定的样式。
    /// </summary>
    /// <param name="styleID">要应用的样式标识符。如果为空字符串，则应用默认样式。</param>
    /// <param name="saveFormatting">是否保留现有的格式设置。</param>
    void ApplyStyle([ComNamespace("MsPowerPoint")] string styleID = "", bool saveFormatting = false);

    /// <summary>
    /// 获取或设置表格的替代文本（用于辅助功能）。
    /// </summary>
    /// <value>表格的替代文本描述。</value>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取或设置表格的标题。
    /// </summary>
    /// <value>表格的标题文本。</value>
    string Title { get; set; }
}