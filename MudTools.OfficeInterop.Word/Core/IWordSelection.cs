//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word Selection 接口，用于操作 Word 文档中的选择区域
/// </summary>
public interface IWordSelection : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取选择区域的文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取选择区域的类型
    /// </summary>
    WdSelectionType Type { get; }

    /// <summary>
    /// 获取或设置选择区域的起始位置
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置选择区域的结束位置
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取选择区域的长度
    /// </summary>
    int Length { get; }

    /// <summary>
    /// 获取父对象（通常是 Document）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取关联的文档
    /// </summary>
    IWordDocument Document { get; }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置是否加粗
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置下划线类型
    /// </summary>
    int Underline { get; set; }

    /// <summary>
    /// 获取或设置文字颜色
    /// </summary>
    WdColor FontColor { get; set; }

    /// <summary>
    /// 获取或设置段落对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置行距
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置首行缩进
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取查找对象
    /// </summary>
    IWordFind Find { get; }

    /// <summary>
    /// 获取范围对象
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 激活选择区域
    /// </summary>
    void Activate();

    /// <summary>
    /// 复制选择区域内容
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切选择区域内容
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到选择区域
    /// </summary>
    void Paste();

    /// <summary>
    /// 删除选择区域内容
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除格式
    /// </summary>
    void ClearFormatting();

    /// <summary>
    /// 插入文本
    /// </summary>
    /// <param name="text">要插入的文本</param>
    void InsertText(string text);

    /// <summary>
    /// 插入段落
    /// </summary>
    void InsertParagraph();

    /// <summary>
    /// 插入换行符
    /// </summary>
    void InsertLineBreak();

    /// <summary>
    /// 插入分页符
    /// </summary>
    void InsertPageBreak();

    /// <summary>
    /// 插入表格
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <returns>表格对象</returns>
    IWordTable InsertTable(int rows, int columns);

    /// <summary>
    /// 向前移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveLeft(int unit = 1, int count = 1);

    /// <summary>
    /// 向后移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveRight(int unit = 1, int count = 1);

    /// <summary>
    /// 向上移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveUp(int unit = 1, int count = 1);

    /// <summary>
    /// 向下移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveDown(int unit = 1, int count = 1);

    /// <summary>
    /// 全选内容
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 取消选择
    /// </summary>
    void Collapse();

    /// <summary>
    /// 扩展选择区域
    /// </summary>
    /// <param name="unit">扩展单位</param>
    /// <param name="count">扩展数量</param>
    void Extend(int unit = 1, int count = 1);

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>是否找到并替换</returns>
    bool FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 设置字符格式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="color">文字颜色</param>
    void SetFont(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0);

    /// <summary>
    /// 设置段落格式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="lineSpacing">行距</param>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    /// <param name="firstLineIndent">首行缩进</param>
    void SetParagraph(int alignment = 0, float lineSpacing = 0, float spaceBefore = 0, float spaceAfter = 0, float firstLineIndent = 0);

    /// <summary>
    /// 获取选择区域的书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    IWordBookmark GetBookmark(string name);

    /// <summary>
    /// 添加书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    IWordBookmark AddBookmark(string name);

    /// <summary>
    /// 获取选择区域的超链接
    /// </summary>
    /// <param name="address">超链接地址</param>
    /// <returns>超链接对象</returns>
    IWordHyperlink AddHyperlink(string address);

    /// <summary>
    /// 获取选择区域内的所有书签
    /// </summary>
    IEnumerable<IWordBookmark> Bookmarks { get; }

    /// <summary>
    /// 获取选择区域内的所有表格
    /// </summary>
    IEnumerable<IWordTable> Tables { get; }

    /// <summary>
    /// 刷新显示
    /// </summary>
    void Refresh();
}