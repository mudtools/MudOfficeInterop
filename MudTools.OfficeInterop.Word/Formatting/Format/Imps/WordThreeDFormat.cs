//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ThreeDFormat 的实现类。
/// </summary>
internal class WordThreeDFormat : IWordThreeDFormat
{
    private MsWord.ThreeDFormat _threeDFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="threeDFormat">原始 COM ThreeDFormat 对象。</param>
    internal WordThreeDFormat(MsWord.ThreeDFormat threeDFormat)
    {
        _threeDFormat = threeDFormat ?? throw new ArgumentNullException(nameof(threeDFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _threeDFormat != null ? new WordApplication(_threeDFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _threeDFormat?.Parent;

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
    public IWordColorFormat? ExtrusionColor =>
        _threeDFormat?.ExtrusionColor != null ? new WordColorFormat(_threeDFormat.ExtrusionColor) : null;

    public IWordColorFormat? ContourColor =>
         _threeDFormat?.ContourColor != null ? new WordColorFormat(_threeDFormat.ContourColor) : null;

    /// <inheritdoc/>
    public MsoExtrusionColorType ExtrusionColorType
    {
        get => _threeDFormat?.ExtrusionColorType != null ? (MsoExtrusionColorType)(int)_threeDFormat?.ExtrusionColorType : MsoExtrusionColorType.msoExtrusionColorAutomatic;
        set
        {
            if (_threeDFormat != null) _threeDFormat.ExtrusionColorType = (MsCore.MsoExtrusionColorType)(int)value;
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

    public bool ProjectText
    {
        get => _threeDFormat?.ProjectText == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.ProjectText = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
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
    public MsoPresetMaterial PresetMaterial
    {
        get => _threeDFormat?.PresetMaterial != null ? (MsoPresetMaterial)(int)_threeDFormat?.PresetMaterial : MsoPresetMaterial.msoPresetMaterialMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetMaterial = (MsCore.MsoPresetMaterial)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoLightRigType PresetLighting
    {
        get => _threeDFormat?.PresetLighting != null ? (MsoLightRigType)(int)_threeDFormat?.PresetLighting : MsoLightRigType.msoLightRigMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetLighting = (MsCore.MsoLightRigType)(int)value;
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
    public MsoPresetCamera PresetCamera
    {
        get => _threeDFormat?.PresetCamera != null ? (MsoPresetCamera)(int)_threeDFormat?.PresetCamera : MsoPresetCamera.msoPresetCameraMixed;
    }

    /// <inheritdoc/>
    public MsoPresetLightingSoftness PresetLightingSoftness
    {
        get => _threeDFormat?.PresetLightingSoftness != null ? (MsoPresetLightingSoftness)(int)_threeDFormat?.PresetLightingSoftness : MsoPresetLightingSoftness.msoPresetLightingSoftnessMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetLightingSoftness = (MsCore.MsoPresetLightingSoftness)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoPresetLightingDirection PresetLightingDirection
    {
        get => _threeDFormat?.PresetLightingDirection != null ? (MsoPresetLightingDirection)(int)_threeDFormat?.PresetLightingDirection : MsoPresetLightingDirection.msoPresetLightingDirectionMixed;
        set
        {
            if (_threeDFormat != null) _threeDFormat.PresetLightingDirection = (MsCore.MsoPresetLightingDirection)(int)value;
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
    public bool PerspectiveEnabled
    {
        get => _threeDFormat?.Perspective == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Perspective = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public MsoBevelType BevelTopType
    {
        get => _threeDFormat?.BevelTopType != null ? (MsoBevelType)(int)_threeDFormat?.BevelTopType : MsoBevelType.msoBevelNone;
        set
        {
            if (_threeDFormat != null) _threeDFormat.BevelTopType = (MsCore.MsoBevelType)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoBevelType BevelBottomType
    {
        get => _threeDFormat?.BevelBottomType != null ? (MsoBevelType)(int)_threeDFormat?.BevelBottomType : MsoBevelType.msoBevelNone;
        set
        {
            if (_threeDFormat != null) _threeDFormat.BevelBottomType = (MsCore.MsoBevelType)(int)value;
        }
    }

    /// <inheritdoc/>
    public float ContourWidth
    {
        get => _threeDFormat?.ContourWidth ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.ContourWidth = value;
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
    public float BevelBottomInset
    {
        get => _threeDFormat?.BevelBottomInset ?? 0f;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.BevelBottomInset = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void SetPresetCamera(MsoPresetCamera presetThreeDFormat)
    {
        _threeDFormat?.SetPresetCamera((MsCore.MsoPresetCamera)(int)presetThreeDFormat);
    }

    /// <inheritdoc/>
    public void SetPresetLighting(MsoLightRigType presetLighting)
    {
        if (_threeDFormat != null)
        {
            _threeDFormat.PresetLighting = (MsCore.MsoLightRigType)(int)presetLighting;
        }
    }

    /// <inheritdoc/>
    public void SetPresetMaterial(MsoPresetMaterial presetMaterial)
    {
        if (_threeDFormat != null)
        {
            _threeDFormat.PresetMaterial = (MsCore.MsoPresetMaterial)(int)presetMaterial;
        }
    }

    /// <inheritdoc/>
    public void SetExtrusionDirection(MsoPresetExtrusionDirection PresetExtrusionDirection)
    {
        _threeDFormat?.SetExtrusionDirection((MsCore.MsoPresetExtrusionDirection)(int)PresetExtrusionDirection);
    }

    /// <inheritdoc/>
    public void SetRotation(float rotationX, float rotationY, float rotationZ)
    {
        if (_threeDFormat != null)
        {
            _threeDFormat.RotationX = rotationX;
            _threeDFormat.RotationY = rotationY;
            _threeDFormat.RotationZ = rotationZ;
        }
    }


    /// <inheritdoc/>
    public void RResetRotationeset()
    {
        _threeDFormat?.ResetRotation();
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_threeDFormat != null)
        {
            _threeDFormat.Visible = MsCore.MsoTriState.msoFalse;
        }
    }


    /// <inheritdoc/>
    public void SetExtrusionColor(int color)
    {
        if (_threeDFormat?.ExtrusionColor != null)
        {
            _threeDFormat.ExtrusionColor.RGB = color;
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
            // 释放 extrusion 颜色对象
            if (_threeDFormat?.ExtrusionColor != null)
            {
                Marshal.ReleaseComObject(_threeDFormat.ExtrusionColor);
            }
            // 释放三维格式对象本身
            if (_threeDFormat != null)
            {
                Marshal.ReleaseComObject(_threeDFormat);
                _threeDFormat = null;
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