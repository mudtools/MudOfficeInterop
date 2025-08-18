//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 文本范围接口
/// </summary>
public interface IPowerPointTextRange : IDisposable
{
    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取文本长度
    /// </summary>
    int Length { get; }

    /// <summary>
    /// 获取起始位置
    /// </summary>
    int Start { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取字体设置
    /// </summary>
    IPowerPointFont Font { get; }

    /// <summary>
    /// 获取段落格式
    /// </summary>
    IPowerPointParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// 获取字符数
    /// </summary>
    int Characters { get; }

    /// <summary>
    /// 获取单词数
    /// </summary>
    int Words { get; }

    /// <summary>
    /// 获取行数
    /// </summary>
    int Lines { get; }

    /// <summary>
    /// 获取段落数
    /// </summary>
    int Paragraphs { get; }

    /// <summary>
    /// 获取句子数
    /// </summary>
    int Sentences { get; }

    /// <summary>
    /// 选择文本范围
    /// </summary>
    void Select();

    /// <summary>
    /// 复制文本范围
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除文本范围
    /// </summary>
    void Delete();

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findWhat">查找内容</param>
    /// <param name="replaceWhat">替换内容</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="wholeWords">是否匹配整个单词</param>
    /// <returns>替换次数</returns>
    int Replace(string findWhat, string replaceWhat, bool matchCase = false, bool wholeWords = false);

    /// <summary>
    /// 插入文本到文本范围
    /// </summary>
    /// <param name="newText">要插入的文本</param>
    /// <param name="start">插入起始位置</param>
    /// <param name="length">插入长度</param>
    /// <returns>新插入的文本范围</returns>
    IPowerPointTextRange InsertAfter(string newText, int start = -1, int length = 0);

    /// <summary>
    /// 在文本范围前插入文本
    /// </summary>
    /// <param name="newText">要插入的文本</param>
    /// <returns>新插入的文本范围</returns>
    IPowerPointTextRange InsertBefore(string newText);

    /// <summary>
    /// 获取指定字符的文本范围
    /// </summary>
    /// <param name="start">起始字符索引</param>
    /// <param name="length">字符长度</param>
    /// <returns>文本范围</returns>
    IPowerPointTextRange CharactersRange(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定单词的文本范围
    /// </summary>
    /// <param name="start">起始单词索引</param>
    /// <param name="length">单词长度</param>
    /// <returns>文本范围</returns>
    IPowerPointTextRange WordsRange(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定行的文本范围
    /// </summary>
    /// <param name="start">起始行索引</param>
    /// <param name="length">行长度</param>
    /// <returns>文本范围</returns>
    IPowerPointTextRange LinesRange(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定段落的文本范围
    /// </summary>
    /// <param name="start">起始段落索引</param>
    /// <param name="length">段落长度</param>
    /// <returns>文本范围</returns>
    IPowerPointTextRange ParagraphsRange(int start = -1, int length = -1);

    /// <summary>
    /// 获取指定句子的文本范围
    /// </summary>
    /// <param name="start">起始句子索引</param>
    /// <param name="length">句子长度</param>
    /// <returns>文本范围</returns>
    IPowerPointTextRange SentencesRange(int start = -1, int length = -1);

    /// <summary>
    /// 设置文本范围的字体格式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="color">字体颜色</param>
    void SetFontFormat(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0);

    /// <summary>
    /// 设置文本范围的段落格式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    /// <param name="lineSpacing">行距</param>
    /// <param name="firstLineIndent">首行缩进</param>
    void SetParagraphFormat(int alignment = 0, float spaceBefore = 0, float spaceAfter = 0, float lineSpacing = 0, float firstLineIndent = 0);

    /// <summary>
    /// 添加超链接到文本范围
    /// </summary>
    /// <param name="address">超链接地址</param>
    /// <returns>超链接对象</returns>
    IPowerPointHyperlink AddHyperlink(string address);

    /// <summary>
    /// 添加动作设置到文本范围
    /// </summary>
    /// <param name="actionType">动作类型</param>
    /// <param name="action">动作设置</param>
    void AddActionSetting(int actionType, object action);

    /// <summary>
    /// 获取文本范围的边界框
    /// </summary>
    /// <param name="left">左边缘</param>
    /// <param name="top">上边缘</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    void GetBoundingBox(out float left, out float top, out float width, out float height);

    /// <summary>
    /// 刷新文本范围显示
    /// </summary>
    void Refresh();
}
