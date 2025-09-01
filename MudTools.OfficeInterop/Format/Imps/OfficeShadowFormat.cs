//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.ShadowFormat 的二次封装实现类。
/// 提供安全访问阴影格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeShadowFormat : IOfficeShadowFormat
{
    private MsCore.ShadowFormat _shadowFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 ShadowFormat 对象。
    /// </summary>
    /// <param name="shadowFormat">原始的 COM ShadowFormat 对象。</param>
    internal OfficeShadowFormat(MsCore.ShadowFormat shadowFormat)
    {
        _shadowFormat = shadowFormat ?? throw new ArgumentNullException(nameof(shadowFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IOfficeColorFormat ForeColor
    {
        get
        {
            if (_shadowFormat?.ForeColor != null)
                return new OfficeColorFormat(_shadowFormat.ForeColor);
            return null;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _shadowFormat?.Transparency ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public float Blur
    {
        get => _shadowFormat?.Blur ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Blur = value;
        }
    }

    /// <inheritdoc/>
    public float OffsetX
    {
        get => _shadowFormat?.OffsetX ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetX = value;
        }
    }

    /// <inheritdoc/>
    public float OffsetY
    {
        get => _shadowFormat?.OffsetY ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetY = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _shadowFormat?.Size ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Size = value;
        }
    }

    /// <inheritdoc/>
    public bool Visible
    {
        get => _shadowFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool RotateWithShape
    {
        get => _shadowFormat?.RotateWithShape == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.RotateWithShape = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }
    /// <inheritdoc/>
    public bool Obscured
    {
        get => _shadowFormat?.Obscured == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Obscured = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }


    /// <inheritdoc/>
    public MsoShadowStyle Style
    {
        get => _shadowFormat?.Style != null ? (MsoShadowStyle)(int)_shadowFormat?.Style : MsoShadowStyle.msoShadowStyleMixed;
        set
        {
            if (_shadowFormat != null) _shadowFormat.Style = (MsCore.MsoShadowStyle)(int)value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Solid(int color)
    {
        if (_shadowFormat?.ForeColor != null)
            _shadowFormat.ForeColor.RGB = color;
    }

    /// <inheritdoc/>
    public void IncrementOffsetX(float offsetX)
    {
        _shadowFormat?.IncrementOffsetX(offsetX);
    }

    /// <inheritdoc/>
    public void IncrementOffsetY(float offsetY)
    {
        _shadowFormat?.IncrementOffsetY(offsetY);
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

        if (disposing && _shadowFormat != null)
        {
            Marshal.ReleaseComObject(_shadowFormat);
            _shadowFormat = null;
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
