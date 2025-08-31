namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// 封装 Microsoft.Office.Core.EffectParameter 的实现类。
/// </summary>
internal class OfficeEffectParameter : IOfficeEffectParameter
{
    private MsCore.EffectParameter _effectParameter;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="effectParameter">原始 COM EffectParameter 对象。</param>
    internal OfficeEffectParameter(MsCore.EffectParameter effectParameter)
    {
        _effectParameter = effectParameter ?? throw new ArgumentNullException(nameof(effectParameter));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string Name => _effectParameter?.Name ?? string.Empty;

    /// <inheritdoc/>
    public object Value
    {
        get => _effectParameter?.Value;
        set
        {
            if (_effectParameter != null)
                _effectParameter.Value = value;
        }
    }
    #endregion

    #region 方法实现
    /// <inheritdoc/>
    public string GetValueAsString()
    {
        if (_effectParameter?.Value == null)
            return string.Empty;

        try
        {
            return _effectParameter.Value.ToString();
        }
        catch
        {
            return string.Empty;
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

        if (disposing && _effectParameter != null)
        {
            Marshal.ReleaseComObject(_effectParameter);
            _effectParameter = null;
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