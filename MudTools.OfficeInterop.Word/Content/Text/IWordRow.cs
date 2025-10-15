//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格行的封装接口。
/// </summary>
public interface IWordRow : IDisposable
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
    /// 获取行索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取行范围。
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取行所在的表格。
    /// </summary>
    IWordTable Table { get; }

    /// <summary>
    /// 获取或设置行高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置行最小高度。
    /// </summary>
    WdRowHeightRule HeightRule { get; set; }

    /// <summary>
    /// 获取或设置行是否允许跨页断行。
    /// </summary>
    bool AllowBreakAcrossPages { get; set; }

    /// <summary>
    /// 获取或设置行是否为标题行。
    /// </summary>
    bool IsHeading { get; set; }

    /// <summary>
    /// 获取行单元格集合。
    /// </summary>
    IWordCells Cells { get; }

    /// <summary>
    /// 获取行边框。
    /// </summary>
    IWordBorders Borders { get; }

    /// <summary>
    /// 获取行底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取行左边界位置。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 选择行。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除行。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取行文本内容。
    /// </summary>
    /// <returns>行文本。</returns>
    string GetText();

    /// <summary>
    /// 设置行文本内容。
    /// </summary>
    /// <param name="text">文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 清除行内容。
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 复制行。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切行。
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到行。
    /// </summary>
    void Paste();

    /// <summary>
    /// 合并行。
    /// </summary>
    /// <param name="mergeTo">要合并到的行。</param>
    void Merge(IWordRow mergeTo);

    /// <summary>
    /// 拆分行。
    /// </summary>
    /// <param name="numRows">拆分后的行数。</param>
    void Split(int numRows);

    /// <summary>
    /// 设置行边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic);

    /// <summary>
    /// 移除行边框。
    /// </summary>
    void RemoveBorders();

    /// <summary>
    /// 设置行底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);

    /// <summary>
    /// 移除行底纹。
    /// </summary>
    void RemoveShading();

    /// <summary>
    /// 获取行字体对象。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取行段落格式对象。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }
}