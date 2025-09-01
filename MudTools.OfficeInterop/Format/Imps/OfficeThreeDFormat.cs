//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.ThreeDFormat 的二次封装实现类。
/// 提供安全访问三维格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeThreeDFormat : IOfficeThreeDFormat
{
    private MsCore.ThreeDFormat _threeDFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 ThreeDFormat 对象。
    /// </summary>
    /// <param name="threeDFormat">原始的 COM ThreeDFormat 对象。</param>
    internal OfficeThreeDFormat(MsCore.ThreeDFormat threeDFormat)
    {
        _threeDFormat = threeDFormat ?? throw new ArgumentNullException(nameof(threeDFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public float Depth
    {
        get => _threeDFormat?.Depth ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Depth = value;
        }
    }

    /// <inheritdoc/>
    public float BevelTopInset
    {
        get => _threeDFormat?.BevelTopInset ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.BevelTopInset = value;
        }
    }

    /// <inheritdoc/>
    public float BevelTopDepth
    {
        get => _threeDFormat?.BevelTopDepth ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.BevelTopDepth = value;
        }
    }

    /// <inheritdoc/>
    public float BevelBottomInset
    {
        get => _threeDFormat?.BevelBottomInset ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.BevelBottomInset = value;
        }
    }

    /// <inheritdoc/>
    public float BevelBottomDepth
    {
        get => _threeDFormat?.BevelBottomDepth ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.BevelBottomDepth = value;
        }
    }

    /// <inheritdoc/>
    public bool Perspective
    {
        get => _threeDFormat?.Perspective == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Perspective = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool Visible
    {
        get => _threeDFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public float RotationX
    {
        get => _threeDFormat?.RotationX ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationX = value;
        }
    }

    /// <inheritdoc/>
    public float RotationY
    {
        get => _threeDFormat?.RotationY ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationY = value;
        }
    }

    /// <inheritdoc/>
    public float RotationZ
    {
        get => _threeDFormat?.RotationZ ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationZ = value;
        }
    }

    /// <inheritdoc/>
    public float Z
    {
        get => _threeDFormat?.Z ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Z = value;
        }
    }


    /// <inheritdoc/>
    public float FieldOfView
    {
        get => _threeDFormat?.FieldOfView ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.FieldOfView = value;
        }
    }

    /// <inheritdoc/>
    public float LightAngle
    {
        get => _threeDFormat?.LightAngle ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.LightAngle = value;
        }
    }

    /// <inheritdoc/>
    public MsoPresetMaterial PresetMaterial
    {
        get => _threeDFormat?.PresetMaterial != null ? (MsoPresetMaterial)(int)_threeDFormat?.PresetMaterial : MsoPresetMaterial.msoPresetMaterialMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetMaterial = (MsCore.MsoPresetMaterial)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoPresetLightingSoftness PresetLightingSoftness
    {
        get => _threeDFormat?.PresetLightingSoftness != null ? (MsoPresetLightingSoftness)(int)_threeDFormat?.PresetLightingSoftness : MsoPresetLightingSoftness.msoLightingNormal;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetLightingSoftness = (MsCore.MsoPresetLightingSoftness)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoPresetLightingDirection PresetLightingDirection
    {
        get => _threeDFormat?.PresetLightingDirection != null ? (MsoPresetLightingDirection)(int)_threeDFormat?.PresetLightingDirection : MsoPresetLightingDirection.msoLightingNone;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetLightingDirection = (MsCore.MsoPresetLightingDirection)(int)value;
        }
    }

    /// <inheritdoc/>
    public IOfficeColorFormat ExtrusionColor
    {
        get
        {
            if (_threeDFormat?.ExtrusionColor != null)
                return new OfficeColorFormat(_threeDFormat.ExtrusionColor);
            return null;
        }
    }

    /// <inheritdoc/>
    public MsoExtrusionColorType ExtrusionColorType
    {
        get => _threeDFormat?.ExtrusionColorType != null ? (MsoExtrusionColorType)(int)_threeDFormat?.ExtrusionColorType : MsoExtrusionColorType.msoExtrusionColorTypeMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.ExtrusionColorType = (MsCore.MsoExtrusionColorType)(int)value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void SetThreeDFormat(MsoPresetThreeDFormat presetCamera)
    {
        _threeDFormat?.SetThreeDFormat((MsCore.MsoPresetThreeDFormat)(int)presetCamera);
    }

    /// <inheritdoc/>
    public void PresetThreeDFormat(MsoPresetExtrusionDirection presetThreeDFormat)
    {
        _threeDFormat?.SetExtrusionDirection((MsCore.MsoPresetExtrusionDirection)(int)presetThreeDFormat);
    }

    /// <inheritdoc/>
    public void IncrementRotationHorizontal(float increment)
    {
        _threeDFormat?.IncrementRotationHorizontal(increment);
    }

    /// <inheritdoc/>
    public void IncrementRotationX(float increment)
    {
        _threeDFormat?.IncrementRotationX(increment);
    }

    /// <inheritdoc/>
    public void IncrementRotationY(float increment)
    {
        _threeDFormat?.IncrementRotationY(increment);
    }

    /// <inheritdoc/>
    public void SetLightRig(MsoPresetCamera presetCamera)
    {
        _threeDFormat?.SetPresetCamera((MsCore.MsoPresetCamera)(int)presetCamera);
    }

    /// <inheritdoc/>
    public void ResetRotation()
    {
        _threeDFormat?.ResetRotation();
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

        if (disposing && _threeDFormat != null)
        {
            Marshal.ReleaseComObject(_threeDFormat);
            _threeDFormat = null;
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