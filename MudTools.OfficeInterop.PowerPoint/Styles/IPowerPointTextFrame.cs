//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 文本框接口
/// </summary>
public interface IPowerPointTextFrame : IDisposable
{
    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取是否有文本
    /// </summary>
    bool HasText { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文本范围
    /// </summary>
    IPowerPointTextRange TextRange { get; }

    /// <summary>
    /// 获取段落格式
    /// </summary>
    IPowerPointParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// 获取字体设置
    /// </summary>
    IPowerPointFont Font { get; }

    /// <summary>
    /// 获取或设置是否自动调整大小
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取或设置垂直锚定位置
    /// </summary>
    int VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置水平锚定位置
    /// </summary>
    int HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本方向
    /// </summary>
    int Orientation { get; set; }

    /// <summary>
    /// 获取或设置边距
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 选择文本框
    /// </summary>
    void Select();

    /// <summary>
    /// 清除文本框内容
    /// </summary>
    void Clear();

    /// <summary>
    /// 添加文本到文本框
    /// </summary>
    /// <param name="text">要添加的文本</param>
    void AddText(string text);

    /// <summary>
    /// 插入文本到指定位置
    /// </summary>
    /// <param name="position">插入位置</param>
    /// <param name="text">要插入的文本</param>
    void InsertText(int position, string text);

    /// <summary>
    /// 删除指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="length">删除长度</param>
    void DeleteText(int start, int length);

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="wholeWords">是否匹配整个单词</param>
    /// <returns>替换次数</returns>
    int ReplaceText(string findText, string replaceText, bool matchCase = false, bool wholeWords = false);

    /// <summary>
    /// 获取指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="length">文本长度</param>
    /// <returns>文本内容</returns>
    string GetTextRange(int start, int length);

    /// <summary>
    /// 设置文本的字体格式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="color">字体颜色</param>
    void SetFontFormat(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0);

    /// <summary>
    /// 设置段落格式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    /// <param name="lineSpacing">行距</param>
    /// <param name="firstLineIndent">首行缩进</param>
    void SetParagraphFormat(int alignment = 0, float spaceBefore = 0, float spaceAfter = 0, float lineSpacing = 0, float firstLineIndent = 0);

    /// <summary>
    /// 自动调整文本框大小
    /// </summary>
    void AutoSizeText();

    /// <summary>
    /// 刷新文本框显示
    /// </summary>
    void Refresh();
}
