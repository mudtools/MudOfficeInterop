namespace MudTools.OfficeInterop.Imp;

/// <summary>
/// 封装 Microsoft.Office.Core.EffectParameters 的实现类。
/// </summary>
internal class OfficeEffectParameters : IOfficeEffectParameters
{
    private MsCore.EffectParameters _effectParameters;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="effectParameters">原始 COM EffectParameters 对象。</param>
    internal OfficeEffectParameters(MsCore.EffectParameters effectParameters)
    {
        _effectParameters = effectParameters ?? throw new ArgumentNullException(nameof(effectParameters));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _effectParameters?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeEffectParameter this[int index]
    {
        get
        {
            if (_effectParameters == null || index < 1 || index > Count)
                return null;

            var parameter = _effectParameters[index];
            return new OfficeEffectParameter(parameter);
        }
    }

    /// <inheritdoc/>
    public IOfficeEffectParameter this[string name]
    {
        get
        {
            if (_effectParameters == null || string.IsNullOrWhiteSpace(name))
                return null;
            var parameter = _effectParameters[name];
            return parameter != null ? new OfficeEffectParameter(parameter) : null;

        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public bool Contains(string name)
    {
        if (_effectParameters == null || string.IsNullOrWhiteSpace(name))
            return false;

        return _effectParameters[name] != null;
    }

    /// <inheritdoc/>
    public List<string> GetAllParameterNames()
    {
        var names = new List<string>();

        if (_effectParameters == null)
            return names;

        for (int i = 1; i <= Count; i++)
        {
            var parameter = _effectParameters[i];
            if (parameter?.Name != null)
            {
                names.Add(parameter.Name);
            }
        }

        return names;
    }

    /// <inheritdoc/>
    public bool SetValue(string name, object value)
    {
        if (_effectParameters == null || string.IsNullOrWhiteSpace(name))
            return false;

        try
        {
            var parameter = _effectParameters[name];
            if (parameter != null)
            {
                parameter.Value = value;
                return true;
            }
            return false;
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public object GetValue(string name)
    {
        if (_effectParameters == null || string.IsNullOrWhiteSpace(name))
            return null;

        try
        {
            var parameter = _effectParameters[name];
            return parameter?.Value;
        }
        catch
        {
            return null;
        }
    }



    /// <inheritdoc/>
    public void CopyTo(IOfficeEffectParameters targetParameters)
    {
        if (_effectParameters == null || targetParameters == null)
            return;

        try
        {
            var targetParams = targetParameters as OfficeEffectParameters;
            if (targetParams?._effectParameters != null)
            {
                for (int i = 1; i <= Count && i <= targetParams.Count; i++)
                {
                    var sourceParam = _effectParameters[i];
                    var targetParam = targetParams._effectParameters[i];

                    if (sourceParam != null && targetParam != null)
                    {
                        targetParam.Value = sourceParam.Value;
                    }
                }
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制参数设置。", ex);
        }
    }
    #endregion

    #region IEnumerable<IOfficeEffectParameter> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeEffectParameter> GetEnumerator()
    {
        if (_effectParameters == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var parameter = _effectParameters[i];
            if (parameter != null)
                yield return new OfficeEffectParameter(parameter);
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
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _effectParameters != null)
        {
            Marshal.ReleaseComObject(_effectParameters);
            _effectParameters = null;
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