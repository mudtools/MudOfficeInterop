//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 文本范围的接口，提供对文本内容及其属性的访问和操作功能
/// </summary>
/// <remarks>
/// 此接口继承自 IEnumerable&lt;IOfficeTextRange2&gt; 和 IDisposable，
/// 支持集合遍历和资源释放功能
/// </remarks>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeTextRange2 : IEnumerable<IOfficeTextRange2>, IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取集合中元素的数量
    /// </summary>
    int Count { get; }

    IOfficeParagraphFormat2 ParagraphFormat { get; }

    /// <summary>
    /// 获取字体格式属性
    /// </summary>
    IOfficeFont2 Font { get; }

    /// <summary>
    /// 根据索引获取子文本范围
    /// </summary>
    /// <param name="index">要获取的文本范围的从零开始的索引</param>
    /// <returns>指定索引处的文本范围，如果索引无效则返回 null</returns>
    IOfficeTextRange2? this[int index] { get; }

    /// <summary>
    /// 根据名称获取子文本范围
    /// </summary>
    /// <param name="name">要获取的文本范围的名称</param>
    /// <returns>具有指定名称的文本范围，如果没有匹配项则返回 null</returns>
    IOfficeTextRange2? this[string name] { get; }

    /// <summary>
    /// 获取段落集合
    /// </summary>
    IOfficeTextRange2 Paragraphs { get; }

    /// <summary>
    /// 获取句子集合
    /// </summary>
    IOfficeTextRange2 Sentences { get; }

    /// <summary>
    /// 获取单词集合
    /// </summary>
    IOfficeTextRange2 Words { get; }

    /// <summary>
    /// 获取字符集合
    /// </summary>
    IOfficeTextRange2 Characters { get; }

    /// <summary>
    /// 获取行集合
    /// </summary>
    IOfficeTextRange2 Lines { get; }

    /// <summary>
    /// 获取运行集合（具有相同字符格式的文本序列）
    /// </summary>
    IOfficeTextRange2 Runs { get; }

    /// <summary>
    /// 获取数学区域集合
    /// </summary>
    IOfficeTextRange2 MathZones { get; }

    /// <summary>
    /// 获取文本长度
    /// </summary>
    int Length { get; }

    /// <summary>
    /// 获取文本范围的起始位置
    /// </summary>
    int Start { get; }

    /// <summary>
    /// 获取文本范围的左边边界坐标
    /// </summary>
    float BoundLeft { get; }

    /// <summary>
    /// 获取文本范围的顶部边界坐标
    /// </summary>
    float BoundTop { get; }

    /// <summary>
    /// 获取文本范围的宽度
    /// </summary>
    float BoundWidth { get; }

    /// <summary>
    /// 获取文本范围的高度
    /// </summary>
    float BoundHeight { get; }

    /// <summary>
    /// 获取或设置语言标识符
    /// </summary>
    MsoLanguageID LanguageID { get; set; }

    /// <summary>
    /// 移除文本前后的空白字符
    /// </summary>
    /// <returns>移除空白字符后的新文本范围</returns>
    IOfficeTextRange2 TrimText();

    /// <summary>
    /// 在当前文本范围之后插入新文本
    /// </summary>
    /// <param name="newText">要插入的文本，默认为空字符串</param>
    /// <returns>包含插入文本的新文本范围</returns>
    IOfficeTextRange2 InsertAfter(string newText = "");

    /// <summary>
    /// 在当前文本范围之前插入新文本
    /// </summary>
    /// <param name="newText">要插入的文本，默认为空字符串</param>
    /// <returns>包含插入文本的新文本范围</returns>
    IOfficeTextRange2 InsertBefore(string newText = "");

    /// <summary>
    /// 插入符号字符
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="charNumber">字符编号</param>
    /// <param name="unicode">是否为 Unicode 字符，默认为 false</param>
    /// <returns>包含插入符号的新文本范围</returns>
    IOfficeTextRange2 InsertSymbol(string fontName, int charNumber, [ConvertTriState] bool unicode = false);

    /// <summary>
    /// 选择文本范围
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切文本范围
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制文本范围
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除文本范围
    /// </summary>
    void Delete();

    /// <summary>
    /// 粘贴剪贴板内容
    /// </summary>
    /// <returns>包含粘贴内容的新文本范围</returns>
    IOfficeTextRange2 Paste();

    /// <summary>
    /// 以特殊格式粘贴剪贴板内容
    /// </summary>
    /// <param name="format">剪贴板格式</param>
    /// <returns>包含粘贴内容的新文本范围</returns>
    IOfficeTextRange2 PasteSpecial(MsoClipboardFormat format);

    /// <summary>
    /// 更改文本大小写
    /// </summary>
    /// <param name="Type">文本大小写更改类型</param>
    void ChangeCase(MsoTextChangeCase Type);

    /// <summary>
    /// 添加句号到文本末尾
    /// </summary>
    void AddPeriods();

    /// <summary>
    /// 移除文本末尾的句号
    /// </summary>
    void RemovePeriods();

    /// <summary>
    /// 查找指定文本
    /// </summary>
    /// <param name="findWhat">要查找的文本</param>
    /// <param name="after">开始查找的位置，默认为 0</param>
    /// <param name="MatchCase">是否区分大小写，默认为 false</param>
    /// <param name="wholeWords">是否只匹配整个单词，默认为 false</param>
    /// <returns>找到的文本范围，未找到则返回 null</returns>
    IOfficeTextRange2 Find(string findWhat, int after = 0, [ConvertTriState] bool MatchCase = false, [ConvertTriState] bool wholeWords = false);

    /// <summary>
    /// 替换指定文本
    /// </summary>
    /// <param name="findWhat">要查找的文本</param>
    /// <param name="replaceWhat">用于替换的文本</param>
    /// <param name="after">开始查找的位置，默认为 0</param>
    /// <param name="matchCase">是否区分大小写，默认为 false</param>
    /// <param name="wholeWords">是否只匹配整个单词，默认为 false</param>
    /// <returns>包含替换文本的新文本范围</returns>
    IOfficeTextRange2 Replace(string findWhat, string replaceWhat, int after = 0, [ConvertTriState] bool matchCase = false, [ConvertTriState] bool wholeWords = false);

    /// <summary>
    /// 将文本运行方向设置为从右到左
    /// </summary>
    void RtlRun();

    /// <summary>
    /// 将文本运行方向设置为从左到右
    /// </summary>
    void LtrRun();

    /// <summary>
    /// 插入图表字段
    /// </summary>
    /// <param name="chartFieldType">图表字段类型</param>
    /// <param name="formula">公式，默认为空字符串</param>
    /// <param name="position">插入位置，默认为 -1（末尾）</param>
    /// <returns>包含插入图表字段的新文本范围</returns>
    IOfficeTextRange2 InsertChartField(MsoChartFieldType chartFieldType, string formula = "", int position = -1);
}