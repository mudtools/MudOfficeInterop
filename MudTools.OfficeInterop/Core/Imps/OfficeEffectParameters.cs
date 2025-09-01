//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

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