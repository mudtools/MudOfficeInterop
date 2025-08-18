//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 文本框实现类
/// </summary>
internal class PowerPointTextFrame : IPowerPointTextFrame
{
    private readonly MsPowerPoint.TextFrame _textFrame;
    private bool _disposedValue;
    private IPowerPointTextRange _textRange;
    private IPowerPointParagraphFormat _paragraphFormat;
    private IPowerPointFont _font;

    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    public string Text
    {
        get => _textFrame.TextRange?.Text ?? string.Empty;
        set
        {
            if (_textFrame.TextRange != null)
            {
                _textFrame.TextRange.Text = value ?? string.Empty;
            }
        }
    }

    /// <summary>
    /// 获取是否有文本
    /// </summary>
    public bool HasText => _textFrame.HasText == MsCore.MsoTriState.msoTrue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _textFrame.Parent;

    /// <summary>
    /// 获取文本范围
    /// </summary>
    public IPowerPointTextRange TextRange
    {
        get
        {
            if (_textRange == null && _textFrame.TextRange != null)
            {
                _textRange = new PowerPointTextRange(_textFrame.TextRange);
            }
            return _textRange;
        }
    }

    /// <summary>
    /// 获取段落格式
    /// </summary>
    public IPowerPointParagraphFormat ParagraphFormat
    {
        get
        {
            if (_paragraphFormat == null && _textFrame.TextRange?.ParagraphFormat != null)
            {
                _paragraphFormat = new PowerPointParagraphFormat(_textFrame.TextRange.ParagraphFormat);
            }
            return _paragraphFormat;
        }
    }

    /// <summary>
    /// 获取字体设置
    /// </summary>
    public IPowerPointFont Font
    {
        get
        {
            if (_font == null && _textFrame.TextRange?.Font != null)
            {
                _font = new PowerPointFont(_textFrame.TextRange.Font);
            }
            return _font;
        }
    }

    /// <summary>
    /// 获取或设置是否自动调整大小
    /// </summary>
    public bool AutoSize
    {
        get => _textFrame.AutoSize == MsPowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
        set => _textFrame.AutoSize = value ? MsPowerPoint.PpAutoSize.ppAutoSizeShapeToFitText : MsPowerPoint.PpAutoSize.ppAutoSizeNone;
    }

    /// <summary>
    /// 获取或设置垂直锚定位置
    /// </summary>
    public int VerticalAnchor
    {
        get => (int)_textFrame.VerticalAnchor;
        set => _textFrame.VerticalAnchor = (MsCore.MsoVerticalAnchor)value;
    }

    /// <summary>
    /// 获取或设置水平锚定位置
    /// </summary>
    public int HorizontalAnchor
    {
        get => (int)_textFrame.HorizontalAnchor;
        set => _textFrame.HorizontalAnchor = (MsCore.MsoHorizontalAnchor)value;
    }

    /// <summary>
    /// 获取或设置文本方向
    /// </summary>
    public int Orientation
    {
        get => (int)_textFrame.Orientation;
        set => _textFrame.Orientation = (MsCore.MsoTextOrientation)value;
    }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    public float MarginLeft
    {
        get => _textFrame.MarginLeft;
        set => _textFrame.MarginLeft = value;
    }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    public float MarginRight
    {
        get => _textFrame.MarginRight;
        set => _textFrame.MarginRight = value;
    }

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    public float MarginTop
    {
        get => _textFrame.MarginTop;
        set => _textFrame.MarginTop = value;
    }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    public float MarginBottom
    {
        get => _textFrame.MarginBottom;
        set => _textFrame.MarginBottom = value;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="textFrame">COM TextFrame 对象</param>
    internal PowerPointTextFrame(MsPowerPoint.TextFrame textFrame)
    {
        _textFrame = textFrame ?? throw new ArgumentNullException(nameof(textFrame));
        _disposedValue = false;
    }

    /// <summary>
    /// 选择文本框
    /// </summary>
    public void Select()
    {
        try
        {
            _textFrame.TextRange?.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select text frame.", ex);
        }
    }

    /// <summary>
    /// 清除文本框内容
    /// </summary>
    public void Clear()
    {
        try
        {
            Text = string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear text frame.", ex);
        }
    }

    /// <summary>
    /// 添加文本到文本框
    /// </summary>
    /// <param name="text">要添加的文本</param>
    public void AddText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        try
        {
            var currentText = Text;
            Text = currentText + text;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add text to text frame.", ex);
        }
    }

    /// <summary>
    /// 插入文本到指定位置
    /// </summary>
    /// <param name="position">插入位置</param>
    /// <param name="text">要插入的文本</param>
    public void InsertText(int position, string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        try
        {
            if (_textFrame.TextRange != null)
            {
                var insertRange = _textFrame.TextRange.Characters(position + 1, 0);
                insertRange.Text = text;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert text into text frame.", ex);
        }
    }

    /// <summary>
    /// 删除指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="length">删除长度</param>
    public void DeleteText(int start, int length)
    {
        if (length <= 0)
            return;

        try
        {
            if (_textFrame.TextRange != null)
            {
                var deleteRange = _textFrame.TextRange.Characters(start + 1, length);
                deleteRange.Text = string.Empty;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete text from text frame.", ex);
        }
    }

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="wholeWords">是否匹配整个单词</param>
    /// <returns>替换次数</returns>
    public int ReplaceText(string findText, string replaceText, bool matchCase = false, bool wholeWords = false)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        try
        {
            var r = _textFrame.TextRange?.Replace(findText, replaceText ?? string.Empty,
                (int)(matchCase ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse));
            return r.Count;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in text frame.", ex);
        }
    }

    /// <summary>
    /// 获取指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="length">文本长度</param>
    /// <returns>文本内容</returns>
    public string GetTextRange(int start, int length)
    {
        try
        {
            if (_textFrame.TextRange != null)
            {
                var range = _textFrame.TextRange.Characters(start + 1, length);
                return range.Text;
            }
            return string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get text range from text frame.", ex);
        }
    }

    /// <summary>
    /// 设置文本的字体格式
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
            if (_textFrame.TextRange?.Font != null)
            {
                var font = _textFrame.TextRange.Font;
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
            throw new InvalidOperationException("Failed to set font format for text frame.", ex);
        }
    }

    /// <summary>
    /// 设置段落格式
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
            if (_textFrame.TextRange?.ParagraphFormat != null)
            {
                var paraFormat = _textFrame.TextRange.ParagraphFormat;
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
            throw new InvalidOperationException("Failed to set paragraph format for text frame.", ex);
        }
    }

    /// <summary>
    /// 自动调整文本框大小
    /// </summary>
    public void AutoSizeText()
    {
        try
        {
            _textFrame.AutoSize = MsPowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to auto-size text frame.", ex);
        }
    }

    /// <summary>
    /// 刷新文本框显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            // 通过重新选择来刷新显示
            _textFrame.TextRange?.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh text frame.", ex);
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
            _textRange?.Dispose();
            _paragraphFormat?.Dispose();
            _font?.Dispose();
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