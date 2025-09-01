//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shape 的实现类。
/// </summary>
internal class WordShape : IWordShape
{
    private MsWord.Shape _shape;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="shape">原始 COM Shape 对象。</param>
    internal WordShape(MsWord.Shape shape)
    {
        _shape = shape ?? throw new ArgumentNullException(nameof(shape));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _shape != null ? new WordApplication(_shape.Application) : null;

    /// <inheritdoc/>
    public object Parent => _shape?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _shape?.Name ?? string.Empty;
        set
        {
            if (_shape != null)
                _shape.Name = value;
        }
    }

    /// <inheritdoc/>
    public MsoShapeType Type => _shape?.Type != null ? (MsoShapeType)(int)_shape.Type : MsoShapeType.msoAutoShape;

    /// <inheritdoc/>
    public float Left
    {
        get => _shape?.Left ?? 0f;
        set
        {
            if (_shape != null)
                _shape.Left = value;
        }
    }

    /// <inheritdoc/>
    public float Top
    {
        get => _shape?.Top ?? 0f;
        set
        {
            if (_shape != null)
                _shape.Top = value;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _shape?.Width ?? 0f;
        set
        {
            if (_shape != null)
                _shape.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _shape?.Height ?? 0f;
        set
        {
            if (_shape != null)
                _shape.Height = value;
        }
    }


    /// <inheritdoc/>
    public WdRelativeHorizontalPosition RelativeHorizontalPosition
    {
        get => (WdRelativeHorizontalPosition)(int)_shape?.RelativeHorizontalPosition;
        set
        {
            if (_shape != null)
                _shape.RelativeHorizontalPosition = (MsWord.WdRelativeHorizontalPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdRelativeVerticalPosition RelativeVerticalPosition
    {
        get => (WdRelativeVerticalPosition)(int)_shape?.RelativeVerticalPosition;
        set
        {
            if (_shape != null)
                _shape.RelativeVerticalPosition = (MsWord.WdRelativeVerticalPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordTextFrame? TextFrame =>
        _shape?.TextFrame != null ? new WordTextFrame(_shape.TextFrame) : null;

    /// <inheritdoc/>
    public IWordFillFormat? Fill =>
        _shape?.Fill != null ? new WordFillFormat(_shape.Fill) : null;

    /// <inheritdoc/>
    public IWordLineFormat? Line =>
        _shape?.Line != null ? new WordLineFormat(_shape.Line) : null;

    /// <inheritdoc/>
    public IWordShadowFormat? Shadow =>
        _shape?.Shadow != null ? new WordShadowFormat(_shape.Shadow) : null;

    /// <inheritdoc/>
    public IWordThreeDFormat? ThreeD =>
        _shape?.ThreeD != null ? new WordThreeDFormat(_shape.ThreeD) : null;

    /// <inheritdoc/>
    public IWordLinkFormat? LinkFormat =>
        _shape?.LinkFormat != null ? new WordLinkFormat(_shape.LinkFormat) : null;

    /// <inheritdoc/>
    public IWordOLEFormat? OLEFormat =>
        _shape?.OLEFormat != null ? new WordOLEFormat(_shape.OLEFormat) : null;

    /// <inheritdoc/>
    public IWordSoftEdgeFormat? SoftEdge =>
        _shape.SoftEdge != null ? new WordSoftEdgeFormat(_shape.SoftEdge) : null;

    /// <inheritdoc/>
    public IWordGlowFormat? Glow =>
         _shape.Glow != null ? new WordGlowFormat(_shape.Glow) : null;

    /// <inheritdoc/>
    public IWordReflectionFormat? Reflection =>
        _shape.Reflection != null ? new WordReflectionFormat(_shape.Reflection) : null;

    /// <inheritdoc/>
    public bool LockAspectRatio
    {
        get => _shape?.LockAspectRatio == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shape != null)
                _shape.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public float Rotation
    {
        get => _shape?.Rotation != null ? _shape.Rotation : 0f;
        set
        {
            if (_shape != null)
                _shape.Rotation = value;
        }
    }

    /// <inheritdoc/>
    public string AlternativeText
    {
        get => _shape?.AlternativeText ?? string.Empty;
        set
        {
            if (_shape != null)
                _shape.AlternativeText = value;
        }
    }

    /// <inheritdoc/>
    public int ZOrderPosition => _shape?.ZOrderPosition ?? 0;

    /// <inheritdoc/>
    public bool IsFloating => _shape?.WrapFormat != null;

    /// <inheritdoc/>
    public bool IsInline => _shape?.WrapFormat == null;

    /// <inheritdoc/>
    public IWordRange Anchor => _shape?.Anchor != null ? new WordRange(_shape.Anchor) : null;

    /// <inheritdoc/>
    public MsWord.Chart Chart => _shape?.Chart;

    /// <inheritdoc/>
    public MsCore.SmartArt SmartArt => _shape?.SmartArt;
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _shape?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _shape?.Select();
    }

    /// <inheritdoc/>
    public void ZOrder(MsoZOrderCmd position)
    {
        _shape?.ZOrder((MsCore.MsoZOrderCmd)(int)position);
    }

    /// <inheritdoc/>
    public void ScaleHeight(float Factor, bool RelativeToOriginalSize, MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft)
    {
        _shape?.ScaleHeight(Factor,
            RelativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
           (MsCore.MsoScaleFrom)(int)Scale);
    }

    /// <inheritdoc/>
    public void ScaleWidth(float Factor, bool RelativeToOriginalSize, MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft)
    {
        _shape?.ScaleWidth(Factor,
           RelativeToOriginalSize ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
          (MsCore.MsoScaleFrom)(int)Scale);
    }

    /// <inheritdoc/>
    public void IncrementRotation(float increment)
    {
        _shape?.IncrementRotation(increment);
    }

    /// <inheritdoc/>
    public void FlipHorizontal()
    {
        _shape?.Flip(MsCore.MsoFlipCmd.msoFlipHorizontal);
    }

    /// <inheritdoc/>
    public void FlipVertical()
    {
        _shape?.Flip(MsCore.MsoFlipCmd.msoFlipVertical);
    }

    /// <inheritdoc/>
    public IWordInlineShape? ConvertToInlineShape()
    {
        var inlineShape = _shape?.ConvertToInlineShape();
        return inlineShape != null ? new WordInlineShape(inlineShape) : null;
    }

    /// <inheritdoc/>
    public IWordFrame? ConvertToFrame()
    {
        var frame = _shape?.ConvertToFrame();
        return frame != null ? new WordFrame(frame) : null;
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
            // 释放锚点范围
            if (_shape?.Anchor != null)
            {
                Marshal.ReleaseComObject(_shape.Anchor);
            }
            // 释放形状对象本身
            if (_shape != null)
            {
                Marshal.ReleaseComObject(_shape);
                _shape = null;
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