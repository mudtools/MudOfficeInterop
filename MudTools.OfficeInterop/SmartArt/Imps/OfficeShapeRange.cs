//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop;
/// <summary>
/// 对 Microsoft.Office.Core.ShapeRange 的二次封装实现类。
/// 提供安全访问形状范围属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeShapeRange : IOfficeShapeRange
{
    private MsCore.ShapeRange _shapeRange;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 ShapeRange 对象。
    /// </summary>
    /// <param name="shapeRange">原始的 COM ShapeRange 对象。</param>
    internal OfficeShapeRange(MsCore.ShapeRange shapeRange)
    {
        _shapeRange = shapeRange ?? throw new ArgumentNullException(nameof(shapeRange));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _shapeRange?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeShape this[int index]
    {
        get
        {
            if (_shapeRange == null || index < 1 || index > Count)
                return null;

            try
            {
                var shape = _shapeRange.Item(index);
                return new OfficeShape(shape);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IOfficeShape this[string name]
    {
        get
        {
            if (_shapeRange == null || string.IsNullOrWhiteSpace(name))
                return null;

            try
            {
                var shape = _shapeRange.Item(name);
                return shape != null ? new OfficeShape(shape) : null;
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
    public float Height
    {
        get => _shapeRange?.Height ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Height = value;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _shapeRange?.Width ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Left
    {
        get => _shapeRange?.Left ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Left = value;
        }
    }

    /// <inheritdoc/>
    public float Top
    {
        get => _shapeRange?.Top ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Top = value;
        }
    }

    /// <inheritdoc/>
    public float Rotation
    {
        get => _shapeRange?.Rotation ?? 0f;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Rotation = value;
        }
    }

    /// <inheritdoc/>
    public bool Connector
    {
        get => _shapeRange?.Connector == MsCore.MsoTriState.msoTrue;
    }

    /// <inheritdoc/>
    public IOfficeConnectorFormat ConnectorFormat
    {
        get
        {
            if (_shapeRange?.Fill != null)
                return new OfficeConnectorFormat(_shapeRange.ConnectorFormat);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeThreeDFormat ThreeD
    {
        get
        {
            if (_shapeRange?.Fill != null)
                return new OfficeThreeDFormat(_shapeRange.ThreeD);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeFillFormat Fill
    {
        get
        {
            if (_shapeRange?.Fill != null)
                return new OfficeFillFormat(_shapeRange.Fill);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeLineFormat Line
    {
        get
        {
            if (_shapeRange?.Line != null)
                return new OfficeLineFormat(_shapeRange.Line);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeShadowFormat Shadow
    {
        get
        {
            if (_shapeRange?.Shadow != null)
                return new OfficeShadowFormat(_shapeRange.Shadow);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeTextEffectFormat TextEffect
    {
        get
        {
            if (_shapeRange?.TextEffect != null)
                return new OfficeTextEffectFormat(_shapeRange.TextEffect);
            return null;
        }
    }
    public IOfficePictureFormat PictureFormat
    {
        get
        {
            if (_shapeRange?.PictureFormat != null)
                return new OfficePictureFormat(_shapeRange.PictureFormat);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeTextFrame TextFrame
    {
        get
        {
            if (_shapeRange?.TextFrame != null)
                return new OfficeTextFrame(_shapeRange.TextFrame);
            return null;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IOfficeShape Group()
    {
        if (_shapeRange == null)
            return null;

        try
        {
            var groupedShape = _shapeRange.Group();
            return new OfficeShape(groupedShape);
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeShapeRange Ungroup()
    {
        if (_shapeRange == null)
            return null;

        try
        {
            var ungroupedRange = _shapeRange.Ungroup();
            return new OfficeShapeRange(ungroupedRange);
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Select(bool replace = true)
    {
        _shapeRange?.Select(replace ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _shapeRange?.Delete();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _shapeRange?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _shapeRange?.Cut();
    }

    /// <inheritdoc/>
    public void Flip(MsoFlipCmd flipCmd)
    {
        _shapeRange?.Flip((MsCore.MsoFlipCmd)(int)flipCmd);
    }

    /// <inheritdoc/>
    public void Align(MsoAlignCmd alignCmd, bool relativeTo = false)
    {
        _shapeRange?.Align((MsCore.MsoAlignCmd)(int)alignCmd, relativeTo ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
    }

    /// <inheritdoc/>
    public void Distribute(MsoDistributeCmd distributeCmd, bool relativeTo = false)
    {
        _shapeRange?.Distribute((MsCore.MsoDistributeCmd)(int)distributeCmd, relativeTo ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse);
    }

    /// <inheritdoc/>
    public void ZOrder(MsoZOrderCmd zOrderCmd)
    {
        _shapeRange?.ZOrder((MsCore.MsoZOrderCmd)(int)zOrderCmd);
    }

    /// <inheritdoc/>
    public void ScaleHeight(float scale, bool scaleWidth, MsoScaleFrom scaleHeight = MsoScaleFrom.msoScaleFromTopLeft)
    {
        _shapeRange?.ScaleHeight(scale,
            scaleWidth ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
           (MsCore.MsoScaleFrom)(int)scaleHeight);
    }

    /// <inheritdoc/>
    public void ScaleWidth(float scale, bool scaleWidth, MsoScaleFrom scaleHeight = MsoScaleFrom.msoScaleFromTopLeft)
    {
        _shapeRange?.ScaleWidth(scale,
            scaleWidth ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse,
            (MsCore.MsoScaleFrom)(int)scaleHeight);
    }

    /// <inheritdoc/>
    public void IncrementLeft(float deltaX)
    {
        _shapeRange?.IncrementLeft(deltaX);
    }

    /// <inheritdoc/>
    public void IncrementTop(float deltaY)
    {
        _shapeRange?.IncrementTop(deltaY);
    }

    /// <inheritdoc/>
    public void IncrementRotation(float increment)
    {
        _shapeRange?.IncrementRotation(increment);
    }

    #endregion

    #region IEnumerable<IOfficeShape> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeShape> GetEnumerator()
    {
        if (_shapeRange == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var shape = _shapeRange.Item(i);
            if (shape != null)
                yield return new OfficeShape(shape);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shapeRange != null)
        {
            Marshal.ReleaseComObject(_shapeRange);
            _shapeRange = null;
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