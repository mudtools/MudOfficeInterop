//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.Shape 的二次封装实现类。
/// 提供安全访问 Shape 属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeShape : IOfficeShape
{
    internal MsCore.Shape _shape;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 OfficeShape 类的新实例
    /// </summary>
    /// <param name="shape">原始的 COM 形状对象</param>
    internal OfficeShape(MsCore.Shape shape)
    {
        _shape = shape ?? throw new ArgumentNullException(nameof(shape));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Id => _shape?.Id ?? -1;

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
    public MsoShapeType Type => _shape?.Type != null ? (MsoShapeType)(int)_shape?.Type : MsoShapeType.msoAutoShape;

    /// <inheritdoc/>
    public string Title => _shape?.Title ?? string.Empty;

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
    public bool Visible
    {
        get => _shape?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shape != null)
                _shape.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public float Left
    {
        get => _shape?.Left ?? 0;
        set
        {
            if (_shape != null)
                _shape.Left = value;
        }
    }

    /// <inheritdoc/>
    public float Top
    {
        get => _shape?.Top ?? 0;
        set
        {
            if (_shape != null)
                _shape.Top = value;
        }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _shape?.Width ?? 0;
        set
        {
            if (_shape != null)
                _shape.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _shape?.Height ?? 0;
        set
        {
            if (_shape != null)
                _shape.Height = value;
        }
    }

    /// <inheritdoc/>
    public int ZOrderPosition => _shape?.ZOrderPosition ?? 0;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _shape?.Delete();
    }

    /// <inheritdoc/>
    public void ZOrder(MsoZOrderCmd ZOrderCmd)
    {
        _shape?.ZOrder((MsCore.MsoZOrderCmd)(int)ZOrderCmd);
    }

    /// <inheritdoc/>
    public void Apply()
    {
        _shape?.Apply();
    }

    /// <inheritdoc/>
    public void Resize(float width, float height)
    {
        if (_shape != null)
        {
            _shape.Width = width;
            _shape.Height = height;
        }
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _shape?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _shape?.Cut();
    }

    /// <inheritdoc/>
    public IOfficeShape Duplicate()
    {
        if (_shape == null)
            return null;

        var duplicatedShape = _shape.Duplicate();
        return duplicatedShape != null ? new OfficeShape(duplicatedShape) : null;
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在处置</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shape != null)
        {
            Marshal.ReleaseComObject(_shape);
            _shape = null;
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