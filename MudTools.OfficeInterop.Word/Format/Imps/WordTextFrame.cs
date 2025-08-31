using MudTools.OfficeInterop.Word.Imps;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TextFrame 的实现类。
/// </summary>
internal class WordTextFrame : IWordTextFrame
{
    private MsWord.TextFrame _textFrame;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="textFrame">原始 COM TextFrame 对象。</param>
    internal WordTextFrame(MsWord.TextFrame textFrame)
    {
        _textFrame = textFrame ?? throw new ArgumentNullException(nameof(textFrame));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _textFrame != null ? new WordApplication(_textFrame.Application) : null;

    /// <inheritdoc/>
    public object Parent => _textFrame?.Parent;

    /// <inheritdoc/>
    public IWordRange TextRange =>
        _textFrame?.TextRange != null ? new WordRange(_textFrame.TextRange) : null;

    /// <inheritdoc/>
    public float MarginLeft
    {
        get => _textFrame?.MarginLeft ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginLeft = value;
        }
    }

    /// <inheritdoc/>
    public float MarginRight
    {
        get => _textFrame?.MarginRight ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginRight = value;
        }
    }

    /// <inheritdoc/>
    public float MarginTop
    {
        get => _textFrame?.MarginTop ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginTop = value;
        }
    }

    /// <inheritdoc/>
    public float MarginBottom
    {
        get => _textFrame?.MarginBottom ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginBottom = value;
        }
    }

    /// <inheritdoc/>
    public MsoHorizontalAnchor HorizontalAnchor
    {
        get => _textFrame?.HorizontalAnchor != null ? (MsoHorizontalAnchor)(int)_textFrame?.HorizontalAnchor : MsoHorizontalAnchor.msoAnchorNone;
        set
        {
            if (_textFrame != null) _textFrame.HorizontalAnchor = (MsCore.MsoHorizontalAnchor)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoVerticalAnchor VerticalAnchor
    {
        get => _textFrame?.VerticalAnchor != null ? (MsoVerticalAnchor)(int)_textFrame?.VerticalAnchor : MsoVerticalAnchor.msoAnchorMiddle;
        set
        {
            if (_textFrame != null) _textFrame.VerticalAnchor = (MsCore.MsoVerticalAnchor)(int)value;
        }
    }

    /// <inheritdoc/>
    public int AutoSize
    {
        get => _textFrame?.AutoSize != null ? _textFrame.AutoSize : 0;
        set
        {
            if (_textFrame != null)
                _textFrame.AutoSize = value;
        }
    }

    /// <inheritdoc/>
    public MsoPathFormat PathFormat
    {
        get => _textFrame?.PathFormat != null ? (MsoPathFormat)(int)_textFrame?.PathFormat : MsoPathFormat.msoPathTypeNone;
        set
        {
            if (_textFrame != null) _textFrame.PathFormat = (MsCore.MsoPathFormat)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordTextFrame? NextFrame =>
        _textFrame?.Next != null ? new WordTextFrame(_textFrame.Next) : null;

    /// <inheritdoc/>
    public IWordTextFrame? PreviousFrame =>
        _textFrame?.Previous != null ? new WordTextFrame(_textFrame.Previous) : null;

    /// <inheritdoc/>
    public IWordShape? ParentShape =>
        _textFrame?.Parent != null ? new WordShape((MsWord.Shape)_textFrame.Parent) : null;

    /// <inheritdoc/>
    public MsoTextOrientation Orientation
    {
        get => _textFrame?.Orientation != null ? (MsoTextOrientation)(int)_textFrame?.Orientation : MsoTextOrientation.msoTextOrientationMixed;
        set
        {
            if (_textFrame != null) _textFrame.Orientation = (MsCore.MsoTextOrientation)(int)value;
        }
    }

    /// <inheritdoc/>
    public float InternalWidth =>
        Math.Max(0, (MarginLeft - MarginRight));

    /// <inheritdoc/>
    public float InternalHeight =>
        Math.Max(0, (MarginTop - MarginBottom));

    /// <inheritdoc/>
    public bool HasText => _textFrame?.HasText == 1;

    /// <inheritdoc/>
    public int CharactersCount => TextRange?.CharactersCount ?? 0;

    /// <inheritdoc/>
    public int ParagraphsCount => TextRange?.ParagraphsCount ?? 0;

    /// <inheritdoc/>
    public IWordFillFormat? Fill =>
        _textFrame?.Parent is MsWord.Shape shape && shape.Fill != null
            ? new WordFillFormat(shape.Fill) : null;

    /// <inheritdoc/>
    public IWordLineFormat? Line =>
        _textFrame?.Parent is MsWord.Shape shape && shape.Line != null
            ? new WordLineFormat(shape.Line) : null;

    /// <inheritdoc/>
    public IWordFont? Font => TextRange?.Font;

    /// <inheritdoc/>
    public IWordParagraphFormat? ParagraphFormat => TextRange?.ParagraphFormat;

    /// <inheritdoc/>
    public bool IsFirstFrame => PreviousFrame == null;

    /// <inheritdoc/>
    public bool IsLastFrame => NextFrame == null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public bool ConnectTo(IWordTextFrame nextTextFrame)
    {
        if (_textFrame == null || nextTextFrame == null)
            return false;

        try
        {
            var targetFrame = (nextTextFrame as WordTextFrame)?._textFrame;
            if (targetFrame != null)
            {
                _textFrame.Next = targetFrame;
                return true;
            }
            return false;
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public void BreakLink()
    {
        _textFrame?.BreakForwardLink();
    }


    /// <inheritdoc/>
    public void SetMargins(float left, float right, float top, float bottom)
    {
        if (_textFrame != null)
        {
            _textFrame.MarginLeft = left;
            _textFrame.MarginRight = right;
            _textFrame.MarginTop = top;
            _textFrame.MarginBottom = bottom;
        }
    }

    /// <inheritdoc/>
    public void SetAlignment(MsoHorizontalAnchor horizontal, MsoVerticalAnchor vertical)
    {
        if (_textFrame != null)
        {
            _textFrame.HorizontalAnchor = (MsCore.MsoHorizontalAnchor)(int)horizontal;
            _textFrame.VerticalAnchor = (MsCore.MsoVerticalAnchor)(int)vertical;
        }
    }

    /// <inheritdoc/>
    public void ClearText()
    {
        if (TextRange != null)
        {
            TextRange.Text = string.Empty;
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordTextFrame targetTextFrame)
    {
        if (_textFrame == null || targetTextFrame == null)
            return;

        try
        {
            // 复制所有基本属性
            targetTextFrame.MarginLeft = this.MarginLeft;
            targetTextFrame.MarginRight = this.MarginRight;
            targetTextFrame.MarginTop = this.MarginTop;
            targetTextFrame.MarginBottom = this.MarginBottom;
            targetTextFrame.HorizontalAnchor = this.HorizontalAnchor;
            targetTextFrame.VerticalAnchor = this.VerticalAnchor;
            targetTextFrame.AutoSize = this.AutoSize;
            targetTextFrame.PathFormat = this.PathFormat;
            targetTextFrame.Orientation = this.Orientation;

            // 复制文本内容
            if (this.HasText && targetTextFrame.TextRange != null)
            {
                targetTextFrame.SetText(this.GetText());
            }

            // 复制字体格式
            if (this.Font != null && targetTextFrame.Font != null)
            {
                targetTextFrame.Font.Name = this.Font.Name;
                targetTextFrame.Font.Size = this.Font.Size;
                targetTextFrame.Font.Bold = this.Font.Bold;
                targetTextFrame.Font.Italic = this.Font.Italic;
                targetTextFrame.Font.Underline = this.Font.Underline;
                targetTextFrame.Font.Color = this.Font.Color;
                targetTextFrame.Font.Superscript = this.Font.Superscript;
                targetTextFrame.Font.Subscript = this.Font.Subscript;
            }

            // 复制段落格式
            if (this.ParagraphFormat != null && targetTextFrame.ParagraphFormat != null)
            {
                targetTextFrame.ParagraphFormat.Alignment = this.ParagraphFormat.Alignment;
                targetTextFrame.ParagraphFormat.FirstLineIndent = this.ParagraphFormat.FirstLineIndent;
                targetTextFrame.ParagraphFormat.LeftIndent = this.ParagraphFormat.LeftIndent;
                targetTextFrame.ParagraphFormat.RightIndent = this.ParagraphFormat.RightIndent;
                targetTextFrame.ParagraphFormat.SpaceBefore = this.ParagraphFormat.SpaceBefore;
                targetTextFrame.ParagraphFormat.SpaceAfter = this.ParagraphFormat.SpaceAfter;
                targetTextFrame.ParagraphFormat.LineSpacingRule = this.ParagraphFormat.LineSpacingRule;
                targetTextFrame.ParagraphFormat.LineSpacing = this.ParagraphFormat.LineSpacing;
            }

            // 复制填充格式（如果父形状存在）
            if (this.Fill != null && targetTextFrame.Fill != null)
            {
                try
                {
                    targetTextFrame.Fill.ForeColor.RGB = this.Fill.ForeColor.RGB;
                    targetTextFrame.Fill.BackColor.RGB = this.Fill.BackColor.RGB;
                    targetTextFrame.Fill.Transparency = this.Fill.Transparency;
                    targetTextFrame.Fill.Visible = this.Fill.Visible;
                }
                catch
                {
                    // 忽略填充格式复制异常
                }
            }

            // 复制边框格式（如果父形状存在）
            if (this.Line != null && targetTextFrame.Line != null)
            {
                try
                {
                    targetTextFrame.Line.ForeColor.RGB = this.Line.ForeColor.RGB;
                    targetTextFrame.Line.Weight = this.Line.Weight;
                    targetTextFrame.Line.Style = this.Line.Style;
                    targetTextFrame.Line.DashStyle = this.Line.DashStyle;
                    targetTextFrame.Line.Visible = this.Line.Visible;
                }
                catch
                {
                    // 忽略边框格式复制异常
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制文本框格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_textFrame != null)
        {
            // 重置所有基本属性为默认值
            _textFrame.MarginLeft = 0.1f;  // Word默认左边距
            _textFrame.MarginRight = 0.1f; // Word默认右边距
            _textFrame.MarginTop = 0.05f;  // Word默认上边距
            _textFrame.MarginBottom = 0.05f; // Word默认下边距
            _textFrame.HorizontalAnchor = MsCore.MsoHorizontalAnchor.msoAnchorNone;
            _textFrame.VerticalAnchor = MsCore.MsoVerticalAnchor.msoAnchorMiddle;
            _textFrame.AutoSize = 0; // wdTextFrameAutoSizeNone
            _textFrame.PathFormat = MsCore.MsoPathFormat.msoPathTypeNone;
            _textFrame.Orientation = MsCore.MsoTextOrientation.msoTextOrientationHorizontal;

            // 清除文本内容
            ClearText();
        }
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return TextRange?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (TextRange != null)
        {
            TextRange.Text = text ?? string.Empty;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放文本范围对象
            if (_textFrame?.TextRange != null)
            {
                Marshal.ReleaseComObject(_textFrame.TextRange);
            }
            // 释放下一个文本框
            if (_textFrame?.Next != null)
            {
                Marshal.ReleaseComObject(_textFrame.Next);
            }
            // 释放上一个文本框
            if (_textFrame?.Previous != null)
            {
                Marshal.ReleaseComObject(_textFrame.Previous);
            }
            // 释放文本框对象本身
            if (_textFrame != null)
            {
                Marshal.ReleaseComObject(_textFrame);
                _textFrame = null;
            }
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}