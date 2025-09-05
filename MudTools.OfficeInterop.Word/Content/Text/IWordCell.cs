//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格单元格的封装接口。
/// </summary>
public interface IWordCell : IDisposable
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
    /// 获取单元格行索引。
    /// </summary>
    int RowIndex { get; }

    /// <summary>
    /// 获取单元格列索引。
    /// </summary>
    int ColumnIndex { get; }

    /// <summary>
    /// 获取单元格范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取单元格所在的行。
    /// </summary>
    IWordRow? Row { get; }

    /// <summary>
    /// 获取单元格所在的列。
    /// </summary>
    IWordColumn? Column { get; }

    /// <summary>
    /// 获取或设置单元格宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置单元格高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置单元格垂直对齐方式。
    /// </summary>
    WdCellVerticalAlignment VerticalAlignment { get; set; }


    /// <summary>
    /// 获取单元格边框。
    /// </summary>
    IWordBorders Borders { get; }

    /// <summary>
    /// 获取单元格底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取单元格首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置单元格首选宽度类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 获取或设置单元格是否适应文本。
    /// </summary>
    bool FitText { get; set; }

    /// <summary>
    /// 获取单元格左边界位置。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取单元格右边界位置。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取单元格上边界位置。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取单元格下边界位置。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取单元格所在的表格。
    /// </summary>
    IWordTables Tables { get; }

    /// <summary>
    /// 选择单元格。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除单元格。
    /// </summary>
    void Delete();

    /// <summary>
    /// 合并单元格。
    /// </summary>
    /// <param name="mergeTo">要合并到的单元格。</param>
    void Merge(IWordCell mergeTo);

    /// <summary>
    /// 拆分单元格。
    /// </summary>
    /// <param name="numRows">行数。</param>
    /// <param name="numColumns">列数。</param>
    void Split(int numRows, int numColumns);

    /// <summary>
    /// 获取单元格文本内容。
    /// </summary>
    /// <returns>单元格文本。</returns>
    string GetText();

    /// <summary>
    /// 设置单元格文本内容。
    /// </summary>
    /// <param name="text">文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 清除单元格内容。
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 复制单元格。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切单元格。
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到单元格。
    /// </summary>
    void Paste();

    /// <summary>
    /// 设置单元格边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic);

    /// <summary>
    /// 移除单元格边框。
    /// </summary>
    void RemoveBorders();

    /// <summary>
    /// 设置单元格底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);

    /// <summary>
    /// 移除单元格底纹。
    /// </summary>
    void RemoveShading();

    /// <summary>
    /// 获取单元格字体对象。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取单元格段落格式对象。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }
}
