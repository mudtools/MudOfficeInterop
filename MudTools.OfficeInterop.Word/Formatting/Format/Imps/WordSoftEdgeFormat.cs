//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Core.SoftEdgeFormat 的实现类。
/// </summary>
internal class WordSoftEdgeFormat : IWordSoftEdgeFormat
{
    private MsWord.SoftEdgeFormat _softEdgeFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="softEdgeFormat">原始 COM SoftEdgeFormat 对象。</param>
    internal WordSoftEdgeFormat(MsWord.SoftEdgeFormat softEdgeFormat)
    {
        _softEdgeFormat = softEdgeFormat ?? throw new ArgumentNullException(nameof(softEdgeFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _softEdgeFormat != null ? new WordApplication(_softEdgeFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _softEdgeFormat?.Parent;

    /// <inheritdoc/>
    public MsoSoftEdgeType Type
    {
        get => _softEdgeFormat?.Type != null ? (MsoSoftEdgeType)(int)_softEdgeFormat?.Type : MsoSoftEdgeType.msoSoftEdgeTypeNone;
        set
        {
            if (_softEdgeFormat != null) _softEdgeFormat.Type = (MsCore.MsoSoftEdgeType)(int)value;
        }
    }

    /// <inheritdoc/>
    public float Radius
    {
        get => _softEdgeFormat?.Radius ?? 0f;
        set
        {
            if (_softEdgeFormat != null && value >= 0)
            {
                _softEdgeFormat.Radius = value;
            }
        }
    }

    /// <inheritdoc/>
    public bool Visible => _softEdgeFormat?.Type != MsCore.MsoSoftEdgeType.msoSoftEdgeTypeNone;

    /// <inheritdoc/>
    public float Size => Radius * 2; // 大小通常为半径的两倍

    /// <inheritdoc/>
    public bool HasSoftEdgeEffect => Type != MsoSoftEdgeType.msoSoftEdgeTypeNone && Radius > 0;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ApplyPreset(MsoSoftEdgeType softEdgeType)
    {
        if (_softEdgeFormat != null)
        {
            try
            {
                _softEdgeFormat.Type = (MsCore.MsoSoftEdgeType)(int)softEdgeType;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"无法应用预设柔化边缘类型 {softEdgeType}。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void SetCustomSoftEdge(float radius, float transparency = 0.5f)
    {
        if (_softEdgeFormat != null)
        {
            if (!ValidateParameters(radius, transparency))
                throw new ArgumentException("柔化边缘参数无效。");

            try
            {
                _softEdgeFormat.Type = MsCore.MsoSoftEdgeType.msoSoftEdgeType1; // 启用自定义柔化边缘
                _softEdgeFormat.Radius = Math.Max(0, radius);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法设置自定义柔化边缘效果。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_softEdgeFormat != null)
        {
            try
            {
                _softEdgeFormat.Type = MsCore.MsoSoftEdgeType.msoSoftEdgeTypeNone;
                _softEdgeFormat.Radius = 0f;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法清除柔化边缘效果。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordSoftEdgeFormat targetSoftEdge)
    {
        if (_softEdgeFormat == null || targetSoftEdge == null)
            return;

        try
        {
            targetSoftEdge.Type = this.Type;
            targetSoftEdge.Radius = this.Radius;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制柔化边缘格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_softEdgeFormat != null)
        {
            try
            {
                _softEdgeFormat.Type = MsCore.MsoSoftEdgeType.msoSoftEdgeTypeNone;
                _softEdgeFormat.Radius = 0f;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法重置柔化边缘格式。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public bool ValidateParameters(float radius, float transparency)
    {
        // 验证半径（非负数）
        if (radius < 0)
            return false;

        // 验证透明度（0.0到1.0之间）
        if (transparency < 0.0f || transparency > 1.0f)
            return false;

        // 半径不能过大（合理的上限）
        if (radius > 100f)
            return false;

        return true;
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

        if (disposing && _softEdgeFormat != null)
        {
            try
            {
                Marshal.ReleaseComObject(_softEdgeFormat);
            }
            catch
            {
                // 忽略释放异常
            }
            _softEdgeFormat = null;
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