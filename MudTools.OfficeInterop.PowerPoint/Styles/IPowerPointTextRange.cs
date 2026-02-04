//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 文本框中的文本范围。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTextRange : IOfficeObject<IPowerPointTextRange, MsPowerPoint.TextRange>, IDisposable
{

    /// <summary>
    /// 获取集合中的文本范围数量。
    /// </summary>
    /// <value>集合中的文本范围数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此文本范围的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此文本范围的父对象。
    /// </summary>
    /// <value>表示此文本范围父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取文本范围的动作设置。
    /// </summary>
    /// <value>表示动作设置的 <see cref="IPowerPointActionSettings"/> 对象。</value>
    IPowerPointActionSettings? ActionSettings { get; }

    /// <summary>
    /// 获取文本范围的起始位置。
    /// </summary>
    /// <value>表示起始位置的整数值。</value>
    int Start { get; }

    /// <summary>
    /// 获取文本范围的长度。
    /// </summary>
    /// <value>表示文本长度的整数值。</value>
    int Length { get; }

    /// <summary>
    /// 获取文本范围的左边界位置（以磅为单位）。
    /// </summary>
    /// <value>表示左边界位置的浮点数。</value>
    float BoundLeft { get; }

    /// <summary>
    /// 获取文本范围的上边界位置（以磅为单位）。
    /// </summary>
    /// <value>表示上边界位置的浮点数。</value>
    float BoundTop { get; }

    /// <summary>
    /// 获取文本范围的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float BoundWidth { get; }

    /// <summary>
    /// 获取文本范围的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float BoundHeight { get; }

    /// <summary>
    /// 获取指定范围的段落集合。
    /// </summary>
    /// <param name="start">起始位置索引。值为-1表示使用默认起始位置。</param>
    /// <param name="length">范围长度。值为-1表示使用默认长度。</param>
    /// <returns>指定范围的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Paragraphs(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定范围的句子集合。
    /// </summary>
    /// <param name="start">起始位置索引。值为-1表示使用默认起始位置。</param>
    /// <param name="length">范围长度。值为-1表示使用默认长度。</param>
    /// <returns>指定范围的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Sentences(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定范围的单词集合。
    /// </summary>
    /// <param name="start">起始位置索引。值为-1表示使用默认起始位置。</param>
    /// <param name="length">范围长度。值为-1表示使用默认长度。</param>
    /// <returns>指定范围的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Words(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定范围的字符集合。
    /// </summary>
    /// <param name="start">起始位置索引。值为-1表示使用默认起始位置。</param>
    /// <param name="length">范围长度。值为-1表示使用默认长度。</param>
    /// <returns>指定范围的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Characters(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定范围的行集合。
    /// </summary>
    /// <param name="start">起始位置索引。值为-1表示使用默认起始位置。</param>
    /// <param name="length">范围长度。值为-1表示使用默认长度。</param>
    /// <returns>指定范围的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Lines(int start = -1, int length = -1);

    /// <summary>
    /// 修剪文本范围的开头和结尾的空格。
    /// </summary>
    /// <returns>修剪后的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? TrimText();

    /// <summary>
    /// 获取或设置文本范围的文本内容。
    /// </summary>
    /// <value>表示文本内容的字符串。</value>
    string? Text { get; set; }

    /// <summary>
    /// 在文本范围后插入新文本。
    /// </summary>
    /// <param name="newText">要插入的新文本。</param>
    /// <returns>插入的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? InsertAfter(string newText = "");

    /// <summary>
    /// 在文本范围前插入新文本。
    /// </summary>
    /// <param name="newText">要插入的新文本。</param>
    /// <returns>插入的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? InsertBefore(string newText = "");

    /// <summary>
    /// 插入日期和时间。
    /// </summary>
    /// <param name="dateTimeFormat">日期时间格式。</param>
    /// <param name="insertAsField">指示是否作为字段插入的布尔值。</param>
    /// <returns>插入的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? InsertDateTime(PpDateTimeFormat dateTimeFormat, [ConvertTriState] bool insertAsField = false);

    /// <summary>
    /// 插入幻灯片编号。
    /// </summary>
    /// <returns>插入的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? InsertSlideNumber();

    /// <summary>
    /// 插入符号。
    /// </summary>
    /// <param name="fontName">字体名称。</param>
    /// <param name="charNumber">字符编码。</param>
    /// <param name="unicode">指示是否为 Unicode 字符的布尔值。</param>
    /// <returns>插入的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? InsertSymbol(string fontName, int charNumber, [ConvertTriState] bool unicode = false);

    /// <summary>
    /// 获取文本范围的字体设置。
    /// </summary>
    /// <value>表示字体设置的 <see cref="IPowerPointFont"/> 对象。</value>
    IPowerPointFont? Font { get; }

    /// <summary>
    /// 获取文本范围的段落格式设置。
    /// </summary>
    /// <value>表示段落格式的 <see cref="IPowerPointParagraphFormat"/> 对象。</value>
    IPowerPointParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取或设置文本范围的缩进级别。
    /// </summary>
    /// <value>表示缩进级别的整数值。</value>
    int IndentLevel { get; set; }

    /// <summary>
    /// 选择此文本范围。
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切此文本范围。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制此文本范围。
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除此文本范围。
    /// </summary>
    void Delete();

    /// <summary>
    /// 粘贴剪贴板内容到此文本范围。
    /// </summary>
    /// <returns>粘贴的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Paste();

    /// <summary>
    /// 更改文本范围的大小写。
    /// </summary>
    /// <param name="type">大小写更改类型。</param>
    void ChangeCase(PpChangeCase type);

    /// <summary>
    /// 为文本范围的每个句子添加句号。
    /// </summary>
    void AddPeriods();

    /// <summary>
    /// 从文本范围的每个句子中移除句号。
    /// </summary>
    void RemovePeriods();

    /// <summary>
    /// 查找文本。
    /// </summary>
    /// <param name="findWhat">要查找的文本。</param>
    /// <param name="after">在指定位置后开始查找。</param>
    /// <param name="matchCase">指示是否区分大小写的布尔值。</param>
    /// <param name="wholeWords">指示是否全字匹配的布尔值。</param>
    /// <returns>找到的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Find(string findWhat, int after = 0, [ConvertTriState] bool matchCase = false, [ConvertTriState] bool wholeWords = false);

    /// <summary>
    /// 替换文本。
    /// </summary>
    /// <param name="findWhat">要查找的文本。</param>
    /// <param name="replaceWhat">要替换的文本。</param>
    /// <param name="after">在指定位置后开始查找。</param>
    /// <param name="matchCase">指示是否区分大小写的布尔值。</param>
    /// <param name="wholeWords">指示是否全字匹配的布尔值。</param>
    /// <returns>替换后的 <see cref="IPowerPointTextRange"/> 对象。</returns>
    IPowerPointTextRange? Replace(string findWhat, string replaceWhat, int after = 0, [ConvertTriState] bool matchCase = false, [ConvertTriState] bool wholeWords = false);

    /// <summary>
    /// 获取旋转后的文本范围边界。
    /// </summary>
    /// <param name="x1">第一个点的 X 坐标。</param>
    /// <param name="y1">第一个点的 Y 坐标。</param>
    /// <param name="x2">第二个点的 X 坐标。</param>
    /// <param name="y2">第二个点的 Y 坐标。</param>
    /// <param name="x3">第三个点的 X 坐标。</param>
    /// <param name="y3">第三个点的 Y 坐标。</param>
    /// <param name="x4">第四个点的 X 坐标。</param>
    /// <param name="y4">第四个点的 Y 坐标。</param>
    void RotatedBounds(out float x1, out float y1, out float x2, out float y2, out float x3, out float y3, out float x4, out float y4);

    /// <summary>
    /// 获取或设置文本范围的语言标识符。
    /// </summary>
    /// <value>表示语言标识符的 <see cref="MsoLanguageID"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLanguageID LanguageID { get; set; }

    /// <summary>
    /// 将文本范围设置为从右到左方向。
    /// </summary>
    void RtlRun();

    /// <summary>
    /// 将文本范围设置为从左到右方向。
    /// </summary>
    void LtrRun();
}