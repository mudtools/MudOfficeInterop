//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Frame 的实现类。
/// </summary>
internal class WordFrame : IWordFrame
{
    private MsWord.Frame _frame;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="frame">原始 COM Frame 对象。</param>
    internal WordFrame(MsWord.Frame frame)
    {
        _frame = frame ?? throw new ArgumentNullException(nameof(frame));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _frame != null ? new WordApplication(_frame.Application) : null;

    /// <inheritdoc/>
    public object Parent => _frame?.Parent;

    /// <inheritdoc/>
    public IWordRange Range =>
        _frame?.Range != null ? new WordRange(_frame.Range) : null;

    /// <inheritdoc/>
    public IWordShading Shading =>
        _frame?.Shading != null ? new WordShading(_frame.Shading) : null;


    /// <inheritdoc/>    
    public float HorizontalPosition
    {
        get => _frame?.HorizontalPosition ?? 0f;
        set
        {
            if (_frame != null)
                _frame.HorizontalPosition = value;
        }
    }

    /// <inheritdoc/>
    public float VerticalPosition
    {
        get => _frame?.VerticalPosition ?? 0f;
        set
        {
            if (_frame != null)
                _frame.VerticalPosition = value;
        }
    }

    /// <inheritdoc/>
    public float HorizontalDistanceFromText
    {
        get => _frame?.HorizontalDistanceFromText ?? 0f;
        set
        {
            if (_frame != null)
                _frame.HorizontalDistanceFromText = value;
        }
    }

    /// <inheritdoc/>
    public float VerticalDistanceFromText
    {
        get => _frame?.VerticalDistanceFromText ?? 0f;
        set
        {
            if (_frame != null)
                _frame.VerticalDistanceFromText = value;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _frame?.Width ?? 0f;
        set
        {
            if (_frame != null)
                _frame.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _frame?.Height ?? 0f;
        set
        {
            if (_frame != null)
                _frame.Height = value;
        }
    }

    /// <inheritdoc/>
    public bool LockAnchor
    {
        get => _frame?.LockAnchor != null ? _frame.LockAnchor : false;
        set
        {
            if (_frame != null)
                _frame.LockAnchor = value;
        }
    }

    /// <inheritdoc/>
    public bool TextWrap
    {
        get => _frame?.TextWrap != null ? _frame.TextWrap : false;
        set
        {
            if (_frame != null)
                _frame.TextWrap = value;
        }
    }

    /// <inheritdoc/>
    public bool HasText => Range?.Text?.Length > 0;

    /// <inheritdoc/>
    public int CharactersCount => Range?.CharactersCount ?? 0;

    /// <inheritdoc/>
    public int ParagraphsCount => Range?.ParagraphsCount ?? 0;

    /// <inheritdoc/>
    public IWordFont Font => Range?.Font;

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat => Range?.ParagraphFormat;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _frame?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _frame?.Select();
    }

    /// <inheritdoc/>
    public void SetDistanceFromText(float horizontal, float vertical)
    {
        if (_frame != null)
        {
            _frame.HorizontalDistanceFromText = horizontal;
            _frame.VerticalDistanceFromText = vertical;
        }
    }

    /// <inheritdoc/>
    public void Resize(float width, float height)
    {
        if (_frame != null)
        {
            _frame.Width = width;
            _frame.Height = height;
        }
    }

    /// <inheritdoc/>
    public void ZOrder(MsoZOrderCmd position)
    {
        if (_frame != null)
        {
            try
            {
                // 如果框架有父形状，可以通过父形状设置Z轴顺序
                var parentShape = Parent as MsWord.Shape;
                parentShape?.ZOrder((MsCore.MsoZOrderCmd)(int)position);
            }
            catch
            {
                // 忽略不支持的ZOrder操作
            }
        }
    }

    /// <inheritdoc/>
    public bool ConnectTo(IWordFrame nextFrame)
    {
        if (_frame == null || nextFrame == null)
            return false;

        try
        {
            var targetFrame = (nextFrame as WordFrame)?._frame;
            if (targetFrame != null)
            {
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
        try
        {
            _frame?.Delete();
        }
        catch
        {
            // 忽略断开连接异常
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordFrame targetFrame)
    {
        if (_frame == null || targetFrame == null)
            return;

        try
        {
            // 复制所有基本属性
            targetFrame.HorizontalPosition = this.HorizontalPosition;
            targetFrame.VerticalPosition = this.VerticalPosition;
            targetFrame.HorizontalDistanceFromText = this.HorizontalDistanceFromText;
            targetFrame.VerticalDistanceFromText = this.VerticalDistanceFromText;
            targetFrame.Width = this.Width;
            targetFrame.Height = this.Height;
            targetFrame.LockAnchor = this.LockAnchor;
            targetFrame.TextWrap = this.TextWrap;

            // 复制文本内容
            if (this.HasText && targetFrame.Range != null)
            {
                targetFrame.SetText(this.GetText());
            }

            // 复制字体格式
            if (this.Font != null && targetFrame.Font != null)
            {
                targetFrame.Font.Name = this.Font.Name;
                targetFrame.Font.Size = this.Font.Size;
                targetFrame.Font.Bold = this.Font.Bold;
                targetFrame.Font.Italic = this.Font.Italic;
                targetFrame.Font.Underline = this.Font.Underline;
                targetFrame.Font.Color = this.Font.Color;
                targetFrame.Font.Superscript = this.Font.Superscript;
                targetFrame.Font.Subscript = this.Font.Subscript;
                targetFrame.Font.Spacing = this.Font.Spacing;
                targetFrame.Font.Scaling = this.Font.Scaling;
                targetFrame.Font.Position = this.Font.Position;
            }

            // 复制段落格式
            if (this.ParagraphFormat != null && targetFrame.ParagraphFormat != null)
            {
                targetFrame.ParagraphFormat.Alignment = this.ParagraphFormat.Alignment;
                targetFrame.ParagraphFormat.FirstLineIndent = this.ParagraphFormat.FirstLineIndent;
                targetFrame.ParagraphFormat.LeftIndent = this.ParagraphFormat.LeftIndent;
                targetFrame.ParagraphFormat.RightIndent = this.ParagraphFormat.RightIndent;
                targetFrame.ParagraphFormat.SpaceBefore = this.ParagraphFormat.SpaceBefore;
                targetFrame.ParagraphFormat.SpaceAfter = this.ParagraphFormat.SpaceAfter;
                targetFrame.ParagraphFormat.LineSpacingRule = this.ParagraphFormat.LineSpacingRule;
                targetFrame.ParagraphFormat.LineSpacing = this.ParagraphFormat.LineSpacing;
                targetFrame.ParagraphFormat.WidowControl = this.ParagraphFormat.WidowControl;
                targetFrame.ParagraphFormat.KeepTogether = this.ParagraphFormat.KeepTogether;
                targetFrame.ParagraphFormat.KeepWithNext = this.ParagraphFormat.KeepWithNext;
            }

            // 复制底纹格式
            if (this.Shading != null && targetFrame.Shading != null)
            {
                try
                {
                    targetFrame.Shading.Texture = this.Shading.Texture;
                    targetFrame.Shading.BackgroundPatternColor = this.Shading.BackgroundPatternColor;
                    targetFrame.Shading.ForegroundPatternColor = this.Shading.ForegroundPatternColor;
                }
                catch
                {
                    // 忽略底纹格式复制异常
                }
            }

            // 复制边框格式（如果可以通过范围获取）
            if (this.Range != null && targetFrame.Range != null)
            {
                try
                {
                    if (this.Range.ParagraphFormat?.Borders != null && targetFrame.Range.ParagraphFormat?.Borders != null)
                    {
                        var sourceBorders = this.Range.ParagraphFormat.Borders;
                        var targetBorders = targetFrame.Range.ParagraphFormat.Borders;
                        targetBorders.ApplyStyle(
                            WdLineStyle.wdLineStyleSingle,
                            WdLineWidth.wdLineWidth050pt,
                            this.Range.ParagraphFormat.Borders[WdBorderType.wdBorderTop]?.Color ?? 0
                        );
                        sourceBorders.Dispose();
                        targetBorders.Dispose();
                    }
                }
                catch
                {
                    // 忽略边框格式复制异常
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制框架格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_frame != null)
        {
            // 重置所有基本属性为默认值
            _frame.HorizontalPosition = 0f;
            _frame.VerticalPosition = 0f;
            _frame.HorizontalDistanceFromText = 0f;
            _frame.VerticalDistanceFromText = 0f;
            _frame.Width = 100f;  // 默认宽度
            _frame.Height = 50f;  // 默认高度
            _frame.LockAnchor = false;
            _frame.TextWrap = false;

            // 清除文本内容
            SetText(string.Empty);

            // 重置字体格式为默认值
            if (Font != null)
            {
                Font.Name = "Calibri";
                Font.Size = 11f;
                Font.Bold = false;
                Font.Italic = false;
                Font.Underline = false;
                Font.Color = 0; // 黑色
            }

            // 重置段落格式为默认值
            if (ParagraphFormat != null)
            {
                ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                ParagraphFormat.FirstLineIndent = 0f;
                ParagraphFormat.LeftIndent = 0f;
                ParagraphFormat.RightIndent = 0f;
                ParagraphFormat.SpaceBefore = 0f;
                ParagraphFormat.SpaceAfter = 0f;
                ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }

            // 清除底纹
            if (Shading != null)
            {
                try
                {
                    Shading.Clear();
                }
                catch
                {
                    // 忽略底纹清除异常
                }
            }
        }
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return Range?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (Range != null)
        {
            Range.Text = text ?? string.Empty;
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
            // 释放范围对象
            if (_frame?.Range != null)
            {
                Marshal.ReleaseComObject(_frame.Range);
            }
            // 释放框架对象本身
            if (_frame != null)
            {
                Marshal.ReleaseComObject(_frame);
                _frame = null;
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