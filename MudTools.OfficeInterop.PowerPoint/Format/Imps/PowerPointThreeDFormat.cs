//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 三维格式实现类
/// </summary>
internal class PowerPointThreeDFormat : IPowerPointThreeDFormat
{
    private readonly MsPowerPoint.ThreeDFormat _threeDFormat;
    private bool _disposedValue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _threeDFormat?.Parent;

    /// <summary>
    /// 获取或设置深度
    /// </summary>
    public float Depth
    {
        get => _threeDFormat?.Depth ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Depth = value;
        }
    }

    /// <summary>
    /// 获取或设置挤出颜色
    /// </summary>
    public int ExtrusionColor
    {
        get => _threeDFormat?.ExtrusionColor.RGB ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.ExtrusionColor.RGB = value;
        }
    }

    /// <summary>
    /// 获取或设置预设光照
    /// </summary>
    public int PresetLighting
    {
        get => _threeDFormat != null ? (int)_threeDFormat.PresetLighting : 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.PresetLighting = (MsCore.MsoLightRigType)value;
        }
    }

    /// <summary>
    /// 获取或设置预设材质
    /// </summary>
    public int PresetMaterial
    {
        get => _threeDFormat != null ? (int)_threeDFormat.PresetMaterial : 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.PresetMaterial = (MsCore.MsoPresetMaterial)value;
        }
    }

    /// <summary>
    /// 获取或设置预设三维格式
    /// </summary>
    public int PresetThreeDFormat
    {
        get => _threeDFormat != null ? (int)_threeDFormat.PresetThreeDFormat : 0;
        set
        {
            // PresetThreeDFormat 是只读属性
        }
    }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _threeDFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置透视效果
    /// </summary>
    public bool Perspective
    {
        get => _threeDFormat?.Perspective == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.Perspective = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }


    /// <summary>
    /// 获取或设置旋转X轴角度
    /// </summary>
    public float RotationX
    {
        get => _threeDFormat?.RotationX ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationX = value;
        }
    }

    /// <summary>
    /// 获取或设置旋转Y轴角度
    /// </summary>
    public float RotationY
    {
        get => _threeDFormat?.RotationY ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationY = value;
        }
    }

    /// <summary>
    /// 获取或设置旋转Z轴角度
    /// </summary>
    public float RotationZ
    {
        get => _threeDFormat?.RotationZ ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.RotationZ = value;
        }
    }


    /// <summary>
    /// 获取或设置光照角度
    /// </summary>
    public float LightAngle
    {
        get => _threeDFormat?.LightAngle ?? 0;
        set
        {
            if (_threeDFormat != null)
                _threeDFormat.LightAngle = value;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="threeDFormat">COM ThreeDFormat 对象</param>
    internal PowerPointThreeDFormat(MsPowerPoint.ThreeDFormat threeDFormat)
    {
        _threeDFormat = threeDFormat; // 可以为 null
        _disposedValue = false;
    }

    /// <summary>
    /// 设置预设三维格式
    /// </summary>
    /// <param name="presetThreeDFormat">预设三维格式</param>
    public void SetThreeDFormat(int presetThreeDFormat)
    {
        try
        {
            _threeDFormat?.SetThreeDFormat((MsCore.MsoPresetThreeDFormat)presetThreeDFormat);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set preset 3D format.", ex);
        }
    }

    /// <summary>
    /// 设置挤出方向
    /// </summary>
    /// <param name="presetExtrusionDirection">预设挤出方向</param>
    public void SetExtrusionDirection(int presetExtrusionDirection)
    {
        try
        {
            _threeDFormat?.SetExtrusionDirection((MsCore.MsoPresetExtrusionDirection)presetExtrusionDirection);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set extrusion direction.", ex);
        }
    }

    /// <summary>
    /// 设置光照效果
    /// </summary>
    /// <param name="presetLighting">预设光照</param>
    /// <param name="lightAngle">光照角度</param>
    public void SetLighting(int presetLighting, float lightAngle = 0)
    {
        try
        {
            if (_threeDFormat != null)
            {
                _threeDFormat.PresetLighting = (MsCore.MsoLightRigType)presetLighting;
                if (lightAngle != 0)
                    _threeDFormat.LightAngle = lightAngle;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set lighting.", ex);
        }
    }

    /// <summary>
    /// 设置材质效果
    /// </summary>
    /// <param name="presetMaterial">预设材质</param>
    public void SetMaterial(int presetMaterial)
    {
        try
        {
            _threeDFormat.PresetMaterial = (MsCore.MsoPresetMaterial)presetMaterial;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set material.", ex);
        }
    }

    /// <summary>
    /// 设置旋转角度
    /// </summary>
    /// <param name="rotationX">X轴旋转角度</param>
    /// <param name="rotationY">Y轴旋转角度</param>
    /// <param name="rotationZ">Z轴旋转角度</param>
    public void SetRotation(float rotationX = 0, float rotationY = 0, float rotationZ = 0)
    {
        try
        {
            if (_threeDFormat != null)
            {
                if (rotationX != 0) _threeDFormat.RotationX = rotationX;
                if (rotationY != 0) _threeDFormat.RotationY = rotationY;
                if (rotationZ != 0) _threeDFormat.RotationZ = rotationZ;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set rotation.", ex);
        }
    }



    /// <summary>
    /// 重置三维格式
    /// </summary>
    public void Reset()
    {
        try
        {
            _threeDFormat.Visible = MsCore.MsoTriState.msoFalse;
            _threeDFormat.Depth = 0;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset 3D format.", ex);
        }
    }

    /// <summary>
    /// 复制三维格式
    /// </summary>
    /// <returns>复制的三维格式对象</returns>
    public IPowerPointThreeDFormat Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法
            throw new NotImplementedException("Duplicating 3D format is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate 3D format.", ex);
        }
    }

    /// <summary>
    /// 应用三维格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    public void ApplyTo(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        try
        {
            // 这需要具体的实现来应用3D格式到形状
            throw new NotImplementedException("Applying 3D format to shape is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply 3D format to shape.", ex);
        }
    }

    /// <summary>
    /// 设置深度和挤出颜色
    /// </summary>
    /// <param name="depth">深度</param>
    /// <param name="extrusionColor">挤出颜色</param>
    public void SetDepthAndColor(float depth, int extrusionColor)
    {
        try
        {
            if (_threeDFormat != null)
            {
                _threeDFormat.Depth = depth;
                _threeDFormat.ExtrusionColor.RGB = extrusionColor;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set depth and color.", ex);
        }
    }

    /// <summary>
    /// 设置透视效果
    /// </summary>
    /// <param name="perspective">是否启用透视</param>
    /// <param name="autoRotation">是否自动旋转</param>
    public void SetPerspective(bool perspective, bool autoRotation = false)
    {
        try
        {
            if (_threeDFormat != null)
            {
                _threeDFormat.Perspective = perspective ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set perspective.", ex);
        }
    }

    /// <summary>
    /// 获取三维信息
    /// </summary>
    /// <returns>三维信息字符串</returns>
    public string GetThreeDInfo()
    {
        try
        {
            return $"3D Visible: {Visible}, Depth: {Depth}, RotationX: {RotationX}, RotationY: {RotationY}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get 3D info.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
