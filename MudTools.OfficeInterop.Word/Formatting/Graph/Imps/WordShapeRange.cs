//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 形状范围对象的封装实现类。
/// </summary>
internal class WordShapeRange : IWordShapeRange
{
    private MsWord.ShapeRange _shapeRange;
    private bool _disposedValue;

    internal WordShapeRange(MsWord.ShapeRange shapeRange)
    {
        _shapeRange = shapeRange ?? throw new ArgumentNullException(nameof(shapeRange));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _shapeRange != null ? new WordApplication(_shapeRange.Application) : null;

    /// <inheritdoc/>
    public object Parent => _shapeRange?.Parent;

    /// <inheritdoc/>
    public int Count => _shapeRange?.Count ?? 0;

    /// <inheritdoc/>
    public IWordShape this[object index]
    {
        get
        {
            if (_shapeRange == null) return null;
            try
            {
                var comShape = _shapeRange[index];
                return comShape != null ? new WordShape(comShape) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public string Name
    {
        get => _shapeRange?.Name ?? string.Empty;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Name = value;
        }
    }

    /// <inheritdoc/>
    public float Left
    {
        get => _shapeRange?.Left ?? 0f;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Left = value;
        }
    }

    /// <inheritdoc/>
    public float Top
    {
        get => _shapeRange?.Top ?? 0f;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Top = value;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _shapeRange?.Width ?? 0f;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _shapeRange?.Height ?? 0f;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Height = value;
        }
    }

    /// <inheritdoc/>
    public WdShapePosition HorizontalFlip => _shapeRange?.HorizontalFlip != null ? (WdShapePosition)(int)_shapeRange?.HorizontalFlip : WdShapePosition.wdShapeLeft;

    /// <inheritdoc/>
    public WdShapePosition VerticalFlip => _shapeRange?.VerticalFlip != null ? (WdShapePosition)(int)_shapeRange?.VerticalFlip : WdShapePosition.wdShapeLeft;

    /// <inheritdoc/>
    public int ZOrderPosition => _shapeRange?.ZOrderPosition ?? 0;

    /// <inheritdoc/>
    public IWordTextFrame? TextFrame => _shapeRange?.TextFrame != null ? new WordTextFrame(_shapeRange.TextFrame) : null;

    /// <inheritdoc/>
    public IWordFillFormat? Fill => _shapeRange?.Fill != null ? new WordFillFormat(_shapeRange.Fill) : null;

    /// <inheritdoc/>
    public IWordLineFormat? Line => _shapeRange?.Line != null ? new WordLineFormat(_shapeRange.Line) : null;

    /// <inheritdoc/>
    public IWordShadowFormat? Shadow => _shapeRange?.Shadow != null ? new WordShadowFormat(_shapeRange.Shadow) : null;

    /// <inheritdoc/>
    public IWordThreeDFormat? ThreeD => _shapeRange?.ThreeD != null ? new WordThreeDFormat(_shapeRange.ThreeD) : null;

    /// <inheritdoc/>
    public IWordAdjustments? Adjustments => _shapeRange?.Adjustments != null ? new WordAdjustments(_shapeRange.Adjustments) : null;

    /// <inheritdoc/>
    public MsoAutoShapeType AutoShapeType
    {
        get => _shapeRange?.AutoShapeType != null ? (MsoAutoShapeType)(int)_shapeRange?.AutoShapeType : MsoAutoShapeType.msoShapeMixed;
        set
        {
            if (_shapeRange != null) _shapeRange.AutoShapeType = (MsCore.MsoAutoShapeType)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordRange Anchor => _shapeRange?.Anchor != null ? new WordRange(_shapeRange.Anchor) : null;

    /// <inheritdoc/>
    public WdRelativeHorizontalPosition RelativeHorizontalPosition
    {
        get => _shapeRange?.RelativeHorizontalPosition != null ? (WdRelativeHorizontalPosition)(int)_shapeRange?.RelativeHorizontalPosition : WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
        set
        {
            if (_shapeRange != null) _shapeRange.RelativeHorizontalPosition = (MsWord.WdRelativeHorizontalPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdRelativeVerticalPosition RelativeVerticalPosition
    {
        get => _shapeRange?.RelativeVerticalPosition != null ? (WdRelativeVerticalPosition)(int)_shapeRange?.RelativeVerticalPosition : WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
        set
        {
            if (_shapeRange != null) _shapeRange.RelativeVerticalPosition = (MsWord.WdRelativeVerticalPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public int LayoutInCell
    {
        get => _shapeRange?.LayoutInCell ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.LayoutInCell = value;
        }
    }


    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _shapeRange?.Delete();
    }

    /// <inheritdoc/>
    public void Align(MsoAlignCmd alignCmd, int relativeTo)
    {
        _shapeRange?.Align((MsCore.MsoAlignCmd)(int)alignCmd, relativeTo);
    }

    /// <inheritdoc/>
    public void Apply()
    {
        _shapeRange?.Apply();
    }

    /// <inheritdoc/>
    public IWordShapeRange? Duplicate()
    {
        var shapeRange = _shapeRange.Duplicate();
        return shapeRange != null ? new WordShapeRange(shapeRange) : null;
    }


    /// <inheritdoc/>
    public void Select()
    {
        _shapeRange?.Select();
    }

    /// <inheritdoc/>
    public void ZOrder(MsoZOrderCmd zOrderCmd)
    {
        _shapeRange?.ZOrder((MsCore.MsoZOrderCmd)(int)zOrderCmd);
    }

    /// <inheritdoc/>
    public IWordShape? Group()
    {
        if (_shapeRange == null) return null;
        try
        {
            var groupedShapeRange = _shapeRange.Group();
            return groupedShapeRange != null ? new WordShape(groupedShapeRange) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordShapeRange? Ungroup()
    {
        if (_shapeRange == null) return null;
        try
        {
            var ungroupedShapeRange = _shapeRange.Ungroup();
            return ungroupedShapeRange != null ? new WordShapeRange(ungroupedShapeRange) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Distribute(MsoDistributeCmd distributeCmd, int relativeTo)
    {
        _shapeRange?.Distribute((MsCore.MsoDistributeCmd)(int)distributeCmd, relativeTo);
    }

    /// <inheritdoc/>
    public void ConvertToInlineShape()
    {
        _shapeRange?.ConvertToInlineShape();
    }

    /// <inheritdoc/>
    public void ConvertToFrame()
    {
        _shapeRange?.ConvertToFrame();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shapeRange != null)
        {
            Marshal.ReleaseComObject(_shapeRange);
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}