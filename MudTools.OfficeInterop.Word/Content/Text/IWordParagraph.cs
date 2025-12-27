//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 段落的封装接口。
/// </summary>
public interface IWordParagraph : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置段落范围。
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取或设置段落对齐方式。
    /// </summary>
    WdParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置段落首行缩进。
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置段落左缩进。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置段落右缩进。
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置段落行距。
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置段落行距规则。
    /// </summary>
    WdLineSpacing LineSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置段落前间距。
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段落后间距。
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置是否保持段落在一起。
    /// </summary>
    int KeepTogether { get; set; }

    /// <summary>
    /// 获取或设置是否与下段保持在一起。
    /// </summary>
    int KeepWithNext { get; set; }

    /// <summary>
    /// 获取或设置段落页面断开控制。
    /// </summary>
    int PageBreakBefore { get; set; }

    /// <summary>
    /// 获取或设置段落大纲级别。
    /// </summary>
    WdOutlineLevel OutlineLevel { get; set; }

    /// <summary>
    /// 获取或设置段落制表符停止点。
    /// </summary>
    IWordTabStops TabStops { get; }

    /// <summary>
    /// 获取或设置段落边框。
    /// </summary>
    IWordBorders Borders { get; }

    /// <summary>
    /// 获取或设置段落底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取段落字符数。
    /// </summary>
    int CharactersCount { get; }

    /// <summary>
    /// 获取段落单词数。
    /// </summary>
    int WordsCount { get; }

    /// <summary>
    /// 获取段落句子数。
    /// </summary>
    int SentencesCount { get; }

    /// <summary>
    /// 获取段落是否为标题。
    /// </summary>
    bool IsHeading { get; }

    /// <summary>
    /// 获取段落是否为空。
    /// </summary>
    bool IsEmpty { get; }

    /// <summary>
    /// 获取段落格式对象。
    /// </summary>
    IWordParagraphFormat Format { get; }

    /// <summary>
    /// 获取段落字体对象。
    /// </summary>
    IWordFont Font { get; }

    IWordParagraph Next(int count = 1);

    IWordParagraph Previous(int count = 1);

    void OutlinePromote();

    void OutlineDemote();

    void OutlineDemoteToBody();
    void Indent();

    void Outdent();

    void SelectNumber();
    void ResetAdvanceTo();

    void ListAdvanceTo(short Level1 = 0, short Level2 = 0, short Level3 = 0, short Level4 = 0, short Level5 = 0, short Level6 = 0, short Level7 = 0, short Level8 = 0, short Level9 = 0);

    void SeparateList();

    void JoinList();

    /// <summary>
    /// 剪切段落。
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到段落。
    /// </summary>
    void Paste();

    /// <summary>
    /// 获取段落文本内容。
    /// </summary>
    /// <returns>段落文本。</returns>
    string GetText();

    /// <summary>
    /// 设置段落文本内容。
    /// </summary>
    /// <param name="text">文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 追加文本到段落末尾。
    /// </summary>
    /// <param name="text">要追加的文本。</param>
    void AppendText(string text);

    /// <summary>
    /// 在段落前插入文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    void InsertBefore(string text);

    /// <summary>
    /// 在段落后插入文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    void InsertAfter(string text);

    /// <summary>
    /// 检查段落是否包含指定文本。
    /// </summary>
    /// <param name="text">要检查的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <returns>是否包含。</returns>
    bool ContainsText(string text, bool matchCase = false);

    /// <summary>
    /// 查找并替换文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>替换次数。</returns>
    int ReplaceText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 设置段落缩进。
    /// </summary>
    /// <param name="leftIndent">左缩进。</param>
    /// <param name="firstLineIndent">首行缩进。</param>
    /// <param name="rightIndent">右缩进。</param>
    void SetIndent(float leftIndent, float firstLineIndent = 0, float rightIndent = 0);

    /// <summary>
    /// 设置段落间距。
    /// </summary>
    /// <param name="before">段前间距。</param>
    /// <param name="after">段后间距。</param>
    void SetSpacing(float before, float after);

    /// <summary>
    /// 设置段落行距。
    /// </summary>
    /// <param name="lineSpacing">行距值。</param>
    /// <param name="rule">行距规则。</param>
    void SetLineSpacing(float lineSpacing, MsWord.WdLineSpacing rule = MsWord.WdLineSpacing.wdLineSpaceSingle);

    /// <summary>
    /// 添加制表符停止点。
    /// </summary>
    /// <param name="position">位置。</param>
    /// <param name="alignment">对齐方式。</param>
    /// <param name="leader">前导符。</param>
    void AddTabStop(float position, MsWord.WdTabAlignment alignment = MsWord.WdTabAlignment.wdAlignTabLeft,
                   MsWord.WdTabLeader leader = MsWord.WdTabLeader.wdTabLeaderSpaces);

    /// <summary>
    /// 清除所有制表符停止点。
    /// </summary>
    void ClearTabStops();

    /// <summary>
    /// 应用项目符号列表。
    /// </summary>
    void ApplyBulletList();

    /// <summary>
    /// 应用编号列表。
    /// </summary>
    void ApplyNumberedList();

    /// <summary>
    /// 设置段落边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
        WdLineWidth lineWidth = WdLineWidth.wdLineWidth050pt,
        WdColor color = WdColor.wdColorGray25);

    /// <summary>
    /// 移除段落边框。
    /// </summary>
    void RemoveBorders();

    /// <summary>
    /// 设置段落底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);

    /// <summary>
    /// 移除段落底纹。
    /// </summary>
    void RemoveShading();
}