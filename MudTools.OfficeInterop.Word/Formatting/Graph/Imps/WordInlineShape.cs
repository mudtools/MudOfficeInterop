//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.InlineShape 的实现类。
/// </summary>
internal class WordInlineShape : IWordInlineShape
{
    private MsWord.InlineShape _inlineShape;
    private bool _disposedValue;
    private float _originalWidth;
    private float _originalHeight;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="inlineShape">原始 COM InlineShape 对象。</param>
    internal WordInlineShape(MsWord.InlineShape inlineShape)
    {
        _inlineShape = inlineShape ?? throw new ArgumentNullException(nameof(inlineShape));
        _disposedValue = false;

        // 保存原始尺寸
        try
        {
            _originalWidth = _inlineShape.Width;
            _originalHeight = _inlineShape.Height;
        }
        catch
        {
            _originalWidth = 100f;
            _originalHeight = 100f;
        }
    }

    #region 属性实现

    /// <inheritdoc/>
    public WdInlineShapeType Type => _inlineShape?.Type != null ? (WdInlineShapeType)(int)_inlineShape?.Type : WdInlineShapeType.wdInlineShapePicture;

    /// <inheritdoc/>
    public IWordTextEffectFormat? TextEffect =>
        _inlineShape?.TextEffect != null ? new WordTextEffectFormat(_inlineShape.TextEffect) : null;

    /// <inheritdoc/>
    public IWordRange Range =>
        _inlineShape?.Range != null ? new WordRange(_inlineShape.Range) : null;

    /// <inheritdoc/>
    public IWordApplication? Application => _inlineShape != null ? new WordApplication(_inlineShape.Application) : null;


    /// <inheritdoc/>
    public object Parent => _inlineShape?.Parent;


    /// <inheritdoc/>
    public float Width
    {
        get => _inlineShape?.Width ?? 0f;
        set
        {
            if (_inlineShape != null)
                _inlineShape.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _inlineShape?.Height ?? 0f;
        set
        {
            if (_inlineShape != null)
                _inlineShape.Height = value;
        }
    }


    /// <inheritdoc/>
    public bool LockAspectRatio
    {
        get => _inlineShape?.LockAspectRatio == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_inlineShape != null)
                _inlineShape.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public IWordOLEFormat? OLEFormat =>
         _inlineShape?.OLEFormat != null ? new WordOLEFormat(_inlineShape.OLEFormat) : null;

    /// <inheritdoc/>
    public IWordLinkFormat? LinkFormat =>
         _inlineShape?.LinkFormat != null ? new WordLinkFormat(_inlineShape.LinkFormat) : null;

    /// <inheritdoc/>
    public IWordField? Field => _inlineShape?.Field != null ? new WordField(_inlineShape.Field) : null;

    /// <inheritdoc/>
    public IWordLineFormat? Line =>
        _inlineShape?.Line != null ? new WordLineFormat(_inlineShape.Line) : null;

    /// <inheritdoc/>
    public IWordFillFormat? Fill =>
        _inlineShape?.Fill != null ? new WordFillFormat(_inlineShape.Fill) : null;

    /// <inheritdoc/>
    public IWordShadowFormat? Shadow =>
        _inlineShape?.Shadow != null ? new WordShadowFormat(_inlineShape.Shadow) : null;

    /// <inheritdoc/>
    public IWordChart? Chart =>
        _inlineShape?.Chart != null ? new WordChart(_inlineShape.Chart) : null;

    /// <inheritdoc/>
    //public IWordSmartArt SmartArt => _inlineShape?.SmartArt;

    /// <inheritdoc/>
    public IWordPictureFormat? PictureFormat =>
         _inlineShape?.PictureFormat != null ? new WordPictureFormat(_inlineShape.PictureFormat) : null;

    /// <inheritdoc/>
    public IWordGroupShapes? GroupItems =>
        _inlineShape?.GroupItems != null ? new WordGroupShapes(_inlineShape.GroupItems) : null;

    /// <inheritdoc/>
    public bool IsPicture => Type == WdInlineShapeType.wdInlineShapePicture;

    /// <inheritdoc/>
    public bool IsOLEObject => Type == WdInlineShapeType.wdInlineShapeOLEControlObject;

    /// <inheritdoc/>
    public bool IsChart => Type == WdInlineShapeType.wdInlineShapeChart;

    /// <inheritdoc/>
    public bool IsFirst => _inlineShape?.Range?.Start == 0;

    /// <inheritdoc/>
    public bool IsLast => _inlineShape?.Range?.End == (_inlineShape?.Range?.Document?.Range()?.End ?? 0);

    /// <inheritdoc/>
    public string AlternativeText
    {
        get => _inlineShape?.AlternativeText ?? string.Empty;
        set
        {
            if (_inlineShape != null)
                _inlineShape.AlternativeText = value;
        }
    }

    /// <inheritdoc/>
    public string Title
    {
        get => _inlineShape?.Title ?? string.Empty;
        set
        {
            if (_inlineShape != null)
                _inlineShape.Title = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _inlineShape?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _inlineShape?.Select();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _inlineShape?.Range?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _inlineShape?.Range?.Cut();
    }

    /// <inheritdoc/>
    public void ScaleSize(float width, float height, bool scale = true)
    {
        if (_inlineShape != null)
        {
            if (scale && LockAspectRatio)
            {
                // 按比例缩放
                float scaleX = width / Width;
                float scaleY = height / Height;
                float scaleRatio = Math.Min(scaleX, scaleY);
                _inlineShape.Width = Width * scaleRatio;
                _inlineShape.Height = Height * scaleRatio;
            }
            else
            {
                _inlineShape.Width = width;
                _inlineShape.Height = height;
            }
        }
    }

    /// <inheritdoc/>
    public IWordShape? ConvertToShape()
    {
        if (_inlineShape == null)
            return null;

        try
        {
            var shape = _inlineShape.ConvertToShape();
            return shape != null ? new WordShape(shape) : null;
        }
        catch (COMException)
        {
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void SetSize(float width, float height)
    {
        if (_inlineShape != null)
        {
            _inlineShape.Width = width;
            _inlineShape.Height = height;
        }
    }

    /// <inheritdoc/>
    public void ResetSize()
    {
        if (_inlineShape != null)
        {
            _inlineShape.Width = _originalWidth;
            _inlineShape.Height = _originalHeight;
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordInlineShape targetInlineShape)
    {
        if (_inlineShape == null || targetInlineShape == null)
            return;

        try
        {
            // 复制基本属性
            targetInlineShape.Width = this.Width;
            targetInlineShape.Height = this.Height;
            targetInlineShape.LockAspectRatio = this.LockAspectRatio;
            targetInlineShape.AlternativeText = this.AlternativeText;
            targetInlineShape.Title = this.Title;

            // 复制线条格式
            if (this.Line != null && targetInlineShape.Line != null)
            {
                try
                {
                    targetInlineShape.Line.ForeColor.RGB = this.Line.ForeColor.RGB;
                    targetInlineShape.Line.BackColor.RGB = this.Line.BackColor.RGB;
                    targetInlineShape.Line.Weight = this.Line.Weight;
                    targetInlineShape.Line.Style = this.Line.Style;
                    targetInlineShape.Line.DashStyle = this.Line.DashStyle;
                    targetInlineShape.Line.Visible = this.Line.Visible;
                    targetInlineShape.Line.Transparency = this.Line.Transparency;
                }
                catch
                {
                    // 忽略线条格式复制异常
                }
            }

            // 复制填充格式
            if (this.Fill != null && targetInlineShape.Fill != null)
            {
                try
                {
                    targetInlineShape.Fill.ForeColor.RGB = this.Fill.ForeColor.RGB;
                    targetInlineShape.Fill.BackColor.RGB = this.Fill.BackColor.RGB;
                    targetInlineShape.Fill.Transparency = this.Fill.Transparency;
                    targetInlineShape.Fill.Visible = this.Fill.Visible;
                }
                catch
                {
                    // 忽略填充格式复制异常
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制内联形状格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_inlineShape != null)
        {
            // 重置基本属性为默认值
            _inlineShape.Width = _originalWidth;
            _inlineShape.Height = _originalHeight;
            _inlineShape.LockAspectRatio = MsCore.MsoTriState.msoTrue; // 默认锁定纵横比
            _inlineShape.AlternativeText = string.Empty;
            _inlineShape.Title = string.Empty;
        }
    }

    /// <inheritdoc/>
    public bool Update()
    {
        if (_inlineShape == null)
            return false;

        try
        {
            if (LinkFormat != null)
            {
                LinkFormat.Update();
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
    public bool BreakLink()
    {
        if (_inlineShape == null)
            return false;

        try
        {
            if (LinkFormat != null)
            {
                LinkFormat.BreakLink();
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
            if (_inlineShape?.Range != null)
            {
                Marshal.ReleaseComObject(_inlineShape.Range);
            }
            // 释放线条格式对象
            if (_inlineShape?.Line != null)
            {
                Marshal.ReleaseComObject(_inlineShape.Line);
            }
            // 释放填充格式对象
            if (_inlineShape?.Fill != null)
            {
                Marshal.ReleaseComObject(_inlineShape.Fill);
            }
            // 释放内联形状对象本身
            if (_inlineShape != null)
            {
                Marshal.ReleaseComObject(_inlineShape);
                _inlineShape = null;
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