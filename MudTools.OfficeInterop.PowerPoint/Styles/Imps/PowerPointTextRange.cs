//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 文本范围实现类
/// </summary>
internal class PowerPointTextRange : IPowerPointTextRange
{
    private readonly MsPowerPoint.TextRange _textRange;
    private bool _disposedValue;
    private IPowerPointFont _font;
    private IPowerPointParagraphFormat _paragraphFormat;

    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    public string Text
    {
        get => _textRange.Text;
        set => _textRange.Text = value ?? string.Empty;
    }

    /// <summary>
    /// 获取文本长度
    /// </summary>
    public int Length => _textRange.Length;

    /// <summary>
    /// 获取起始位置
    /// </summary>
    public int Start => _textRange.Start;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _textRange.Parent;

    /// <summary>
    /// 获取字体设置
    /// </summary>
    public IPowerPointFont Font
    {
        get
        {
            if (_font == null && _textRange.Font != null)
            {
                _font = new PowerPointFont(_textRange.Font);
            }
            return _font;
        }
    }

    /// <summary>
    /// 获取段落格式
    /// </summary>
    public IPowerPointParagraphFormat ParagraphFormat
    {
        get
        {
            if (_paragraphFormat == null && _textRange.ParagraphFormat != null)
            {
                _paragraphFormat = new PowerPointParagraphFormat(_textRange.ParagraphFormat);
            }
            return _paragraphFormat;
        }
    }

    /// <summary>
    /// 获取字符数
    /// </summary>
    public int Characters => _textRange.Characters().Count;

    /// <summary>
    /// 获取单词数
    /// </summary>
    public int Words => _textRange.Words().Count;

    /// <summary>
    /// 获取行数
    /// </summary>
    public int Lines => _textRange.Lines().Count;

    /// <summary>
    /// 获取段落数
    /// </summary>
    public int Paragraphs => _textRange.Paragraphs().Count;

    /// <summary>
    /// 获取句子数
    /// </summary>
    public int Sentences => _textRange.Sentences().Count;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="textRange">COM TextRange 对象</param>
    internal PowerPointTextRange(MsPowerPoint.TextRange textRange)
    {
        _textRange = textRange ?? throw new ArgumentNullException(nameof(textRange));
        _disposedValue = false;
    }

    /// <summary>
    /// 选择文本范围
    /// </summary>
    public void Select()
    {
        try
        {
            _textRange.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select text range.", ex);
        }
    }

    /// <summary>
    /// 复制文本范围
    /// </summary>
    public void Copy()
    {
        try
        {
            _textRange.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy text range.", ex);
        }
    }

    /// <summary>
    /// 删除文本范围
    /// </summary>
    public void Delete()
    {
        try
        {
            _textRange.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete text range.", ex);
        }
    }

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findWhat">查找内容</param>
    /// <param name="replaceWhat">替换内容</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="wholeWords">是否匹配整个单词</param>
    /// <returns>替换次数</returns>
    public int Replace(string findWhat, string replaceWhat, bool matchCase = false, bool wholeWords = false)
    {
        if (string.IsNullOrEmpty(findWhat))
            return 0;

        try
        {
            var r = _textRange.Replace(findWhat, replaceWhat ?? string.Empty,
                (int)(matchCase ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse));
            return r.Count;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in text range.", ex);
        }
    }

    /// <summary>
    /// 插入文本到文本范围
    /// </summary>
    /// <param name="newText">要插入的文本</param>
    /// <param name="start">插入起始位置</param>
    /// <param name="length">插入长度</param>
    /// <returns>新插入的文本范围</returns>
    public IPowerPointTextRange InsertAfter(string newText, int start = -1, int length = 0)
    {
        try
        {
            MsPowerPoint.TextRange insertedRange;
            if (start >= 0)
            {
                var targetRange = _textRange.Characters(start + 1, length > 0 ? length : 0);
                insertedRange = targetRange.InsertAfter(newText ?? string.Empty);
            }
            else
            {
                insertedRange = _textRange.InsertAfter(newText ?? string.Empty);
            }
            return new PowerPointTextRange(insertedRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert text after text range.", ex);
        }
    }

    /// <summary>
    /// 在文本范围前插入文本
    /// </summary>
    /// <param name="newText">要插入的文本</param>
    /// <returns>新插入的文本范围</returns>
    public IPowerPointTextRange InsertBefore(string newText)
    {
        try
        {
            var insertedRange = _textRange.InsertBefore(newText ?? string.Empty);
            return new PowerPointTextRange(insertedRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert text before text range.", ex);
        }
    }

    /// <summary>
    /// 获取指定字符的文本范围
    /// </summary>
    /// <param name="start">起始字符索引</param>
    /// <param name="length">字符长度</param>
    /// <returns>文本范围</returns>
    public IPowerPointTextRange CharactersRange(int start = -1, int length = -1)
    {
        try
        {
            MsPowerPoint.TextRange charRange;
            if (start >= 0)
            {
                charRange = _textRange.Characters(start + 1, length > 0 ? length : 1);
            }
            else
            {
                charRange = _textRange.Characters();
            }
            return new PowerPointTextRange(charRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get characters range.", ex);
        }
    }

    /// <summary>
    /// 获取指定单词的文本范围
    /// </summary>
    /// <param name="start">起始单词索引</param>
    /// <param name="length">单词长度</param>
    /// <returns>文本范围</returns>
    public IPowerPointTextRange WordsRange(int start = -1, int length = -1)
    {
        try
        {
            MsPowerPoint.TextRange wordRange;
            if (start >= 0)
            {
                wordRange = _textRange.Words(start + 1, length > 0 ? length : 1);
            }
            else
            {
                wordRange = _textRange.Words();
            }
            return new PowerPointTextRange(wordRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get words range.", ex);
        }
    }

    /// <summary>
    /// 获取指定行的文本范围
    /// </summary>
    /// <param name="start">起始行索引</param>
    /// <param name="length">行长度</param>
    /// <returns>文本范围</returns>
    public IPowerPointTextRange LinesRange(int start = -1, int length = -1)
    {
        try
        {
            MsPowerPoint.TextRange lineRange;
            if (start >= 0)
            {
                lineRange = _textRange.Lines(start + 1, length > 0 ? length : 1);
            }
            else
            {
                lineRange = _textRange.Lines();
            }
            return new PowerPointTextRange(lineRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get lines range.", ex);
        }
    }

    /// <summary>
    /// 获取指定段落的文本范围
    /// </summary>
    /// <param name="start">起始段落索引</param>
    /// <param name="length">段落长度</param>
    /// <returns>文本范围</returns>
    public IPowerPointTextRange ParagraphsRange(int start = -1, int length = -1)
    {
        try
        {
            MsPowerPoint.TextRange paraRange;
            if (start >= 0)
            {
                paraRange = _textRange.Paragraphs(start + 1, length > 0 ? length : 1);
            }
            else
            {
                paraRange = _textRange.Paragraphs();
            }
            return new PowerPointTextRange(paraRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get paragraphs range.", ex);
        }
    }

    /// <summary>
    /// 获取指定句子的文本范围
    /// </summary>
    /// <param name="start">起始句子索引</param>
    /// <param name="length">句子长度</param>
    /// <returns>文本范围</returns>
    public IPowerPointTextRange SentencesRange(int start = -1, int length = -1)
    {
        try
        {
            MsPowerPoint.TextRange sentenceRange;
            if (start >= 0)
            {
                sentenceRange = _textRange.Sentences(start + 1, length > 0 ? length : 1);
            }
            else
            {
                sentenceRange = _textRange.Sentences();
            }
            return new PowerPointTextRange(sentenceRange);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get sentences range.", ex);
        }
    }

    /// <summary>
    /// 设置文本范围的字体格式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="color">字体颜色</param>
    public void SetFontFormat(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0)
    {
        try
        {
            if (_textRange.Font != null)
            {
                var font = _textRange.Font;
                if (!string.IsNullOrEmpty(fontName))
                    font.Name = fontName;
                if (fontSize > 0)
                    font.Size = fontSize;
                font.Bold = bold ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                font.Italic = italic ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                if (underline >= 0)
                    font.Underline = (MsCore.MsoTriState)underline;
                if (color >= 0)
                    font.Color.RGB = color;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set font format for text range.", ex);
        }
    }

    /// <summary>
    /// 设置文本范围的段落格式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    /// <param name="lineSpacing">行距</param>
    /// <param name="firstLineIndent">首行缩进</param>
    public void SetParagraphFormat(int alignment = 0, float spaceBefore = 0, float spaceAfter = 0, float lineSpacing = 0, float firstLineIndent = 0)
    {
        try
        {
            if (_textRange.ParagraphFormat != null)
            {
                var paraFormat = _textRange.ParagraphFormat;
                if (alignment >= 0)
                    paraFormat.Alignment = (MsPowerPoint.PpParagraphAlignment)alignment;
                if (spaceBefore >= 0)
                    paraFormat.SpaceBefore = spaceBefore;
                if (spaceAfter >= 0)
                    paraFormat.SpaceAfter = spaceAfter;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set paragraph format for text range.", ex);
        }
    }

    /// <summary>
    /// 添加超链接到文本范围
    /// </summary>
    /// <param name="address">超链接地址</param>
    /// <returns>超链接对象</returns>
    public IPowerPointHyperlink AddHyperlink(string address)
    {
        if (string.IsNullOrEmpty(address))
            throw new ArgumentException("Hyperlink address cannot be null or empty.", nameof(address));

        try
        {
            var hyperlink = _textRange.ActionSettings[MsPowerPoint.PpMouseActivation.ppMouseClick].Hyperlink;
            hyperlink.Address = address;
            return new PowerPointHyperlink(hyperlink);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add hyperlink to text range.", ex);
        }
    }

    /// <summary>
    /// 添加动作设置到文本范围
    /// </summary>
    /// <param name="actionType">动作类型</param>
    /// <param name="action">动作设置</param>
    public void AddActionSetting(int actionType, object action)
    {
        try
        {
            var actionSetting = _textRange.ActionSettings[(MsPowerPoint.PpMouseActivation)actionType];
            // 根据具体动作类型设置相应的属性
            throw new NotImplementedException("Action setting implementation is not complete.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add action setting to text range.", ex);
        }
    }

    /// <summary>
    /// 获取文本范围的边界框
    /// </summary>
    /// <param name="left">左边缘</param>
    /// <param name="top">上边缘</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    public void GetBoundingBox(out float left, out float top, out float width, out float height)
    {
        try
        {
            // PowerPoint TextRange 没有直接的边界框方法，需要通过其他方式获取
            left = 0;
            top = 0;
            width = 0;
            height = 0;

            // 这里需要更复杂的实现来获取实际边界框
            throw new NotImplementedException("Getting bounding box is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get bounding box for text range.", ex);
        }
    }

    /// <summary>
    /// 刷新文本范围显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            _textRange.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh text range.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _font?.Dispose();
            _paragraphFormat?.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
