//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格列的封装接口。
/// </summary>
public interface IWordColumn : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取列索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取列所在的表格。
    /// </summary>
    IWordTable Table { get; }

    /// <summary>
    /// 获取或设置列宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取列单元格集合。
    /// </summary>
    IWordCells Cells { get; }

    /// <summary>
    /// 获取列边框。
    /// </summary>
    IWordBorders Borders { get; }

    /// <summary>
    /// 获取列底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取或设置列首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置列首选宽度类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }


    /// <summary>
    /// 选择列。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除列。
    /// </summary>
    void Delete();

    /// <summary>
    /// 设置列边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic);

    /// <summary>
    /// 移除列边框。
    /// </summary>
    void RemoveBorders();

    /// <summary>
    /// 设置列底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);

    /// <summary>
    /// 移除列底纹。
    /// </summary>
    void RemoveShading();
}